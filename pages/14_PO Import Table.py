"""
Streamlit app for FULL PIPELINE (FINAL) with multi-row pack detection
- Supports multiple packing-plan Excel uploads (DETAIL sheet expected)
- Supports lookup Excel upload
- Preserves all original pipeline logic and debug info
- Outputs: combined Rekap Excel, Unmatched Excel, per-PO exports (zipped), and displayable previews

Usage: run with `streamlit run streamlit_app_full_pipeline.py`
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
import tempfile
import os
import re
from datetime import datetime

st.set_page_config(page_title="Packing Plan — Full Pipeline", layout="wide")

st.title("Packing Plan — Full Pipeline (with Multi-row Detection)")
st.markdown("Upload packing-plan files (Excel) and a lookup file. The app will run the full pipeline and provide downloadable outputs.")

# ---------------- Configuration panel ----------------
with st.sidebar.form(key='cfg'):
    st.header("Config & Options")
    debug = st.checkbox("DEBUG (show detailed tables)", value=True)
    normalize_pr_strip = st.checkbox("Normalize Packing Rule No. (.str.strip())", value=True)
    auto_group_ffill = st.checkbox("Group rows without Packing Rule No. (ffill)", value=True)
    run_button = st.form_submit_button("Run pipeline")

# ---------------- File uploads ----------------
st.subheader("Step 1 — Upload files")
uploaded_packs = st.file_uploader("Upload one or more packing-plan Excel files (will read sheet 'Detail')", type=["xls","xlsx"], accept_multiple_files=True)
uploaded_lookup = st.file_uploader("Upload lookup Excel (FlexView)", type=["xls","xlsx"]) 

# helper: read Excel file bytes into pandas (return DataFrame or None)
def safe_read_excel_bytes(bytes_io, sheet_name=None, header=None, dtype=str):
    try:
        return pd.read_excel(bytes_io, sheet_name=sheet_name, header=header, dtype=dtype)
    except Exception as e:
        return None

# Placeholders for outputs
combined_bytes = None
unmatched_bytes = None
per_po_zip_bytes = None
run_logs = []

# ---------------- Core pipeline functions (copied from user's pipeline, preserved) ----------------
# Due to length, key helper functions and pipeline are included here (kept faithful to original script)

# --- helpers ---
def norm_size_preserve_dash(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def find_header_row(df: pd.DataFrame, keywords):
    for idx in df.index:
        row = " ".join(df.iloc[idx].astype(str).str.lower().tolist())
        score = sum(1 for k in keywords if k in row)
        if score >= 2:
            return idx
    return None


def find_col_precise(cols, candidates, data, prefer_integer=False):
    cols_lc = [str(c) for c in cols]
    patterns = [re.compile(r"(^|\W)"+re.escape(c.lower())+r"($|\W)") for c in candidates]
    matches = []
    for i,c in enumerate(cols_lc):
        for p in patterns:
            if p.search(c.lower()):
                matches.append(cols[i])
                break
    if not matches:
        for i,c in enumerate(cols_lc):
            for cand in candidates:
                if cand in c.lower():
                    matches.append(cols[i])
                    break
    if not matches:
        return None
    if len(matches) == 1:
        return matches[0]

    best = None
    best_score = -1
    for col in matches:
        series = data[col].dropna().astype(str).str.strip()
        if len(series) == 0:
            score = 0
        else:
            num_prop = series.str.match(r'^-?\d+(\.\d+)?$').mean() if len(series)>0 else 0
            int_prop = series.str.match(r'^\d+$').mean() if len(series)>0 else 0
            uniq_prop = series.nunique() / max(1, len(series))
            nonnull_prop = series.notna().mean()
            if prefer_integer:
                score = int_prop * 0.7 + num_prop * 0.2 + uniq_prop * 0.1
            else:
                score = nonnull_prop * 0.55 + uniq_prop * 0.35 + num_prop * 0.1
        if score > best_score:
            best_score = score
            best = col
    return best


# --- extraction function (reads sheet 'Detail') ---
def extract_package_detail_bytes(file_bytes, filename='<memory>', debug=False):
    x = safe_read_excel_bytes(io.BytesIO(file_bytes), sheet_name='Detail', header=None, dtype=str)
    if x is None:
        if debug: run_logs.append(f"[ERR] {filename}: cannot open sheet 'Detail'")
        return pd.DataFrame()

    df = x.copy()
    pkg_rows = df.index[df.apply(lambda r: r.astype(str).str.contains("package detail", case=False).any(), axis=1)].tolist()
    header_row = None

    if pkg_rows:
        for offset in (1,2,3):
            cand = pkg_rows[0] + offset
            if cand < len(df):
                rowtxt = " ".join(df.iloc[cand].astype(str).str.lower().tolist())
                if any(k in rowtxt for k in ["range","serial","buyer","po","pkg","gross","qty","manufacturing","size"]):
                    header_row = cand
                    if debug: run_logs.append(f"[INFO] {filename}: header at {cand} (after PACKAGE DETAIL)")
                    break

    if header_row is None:
        header_keywords = ["range","serial","buyer","po","pkg","gross","qty per pkg","manufacturing size","buyer item","size"]
        header_row = find_header_row(df, header_keywords)
        if debug and header_row is not None:
            run_logs.append(f"[INFO] {filename}: header fallback at {header_row}")

    if header_row is None:
        if debug: run_logs.append(f"[WARN] {filename}: header not found. skip.")
        return pd.DataFrame()

    header = df.iloc[header_row].fillna("").astype(str).tolist()
    data = df.iloc[header_row+1:].copy()
    data.columns = header

    candidate_patterns = [r'pkg', r'inner', r'qty', r'item', r'item qty', r'qty per', r'serial', r'\bfrom\b', r'\bto\b']
    fill_cols = []
    for c in data.columns:
        lc = str(c).lower()
        if any(re.search(p, lc) for p in candidate_patterns):
            fill_cols.append(c)

    if fill_cols:
        data[fill_cols] = data[fill_cols].ffill(axis=0).bfill(axis=0)

    data = data.dropna(how='all').copy()
    if data.empty:
        if debug: run_logs.append(f"[WARN] {filename}: no data after header/ffill.")
        return pd.DataFrame()

    cols = data.columns.tolist()

    mapping_candidates = {
        "PO #": (["po number","po #","po no","^po$","purchase order"], False),
        "Range": (["^range$","^range\\b","range"], False),
        "Buyer Item #": (["buyer item","buyer item #","article","sku","style"], False),
        "Manufacturing Size": (["^manufacturing size$","manufacturing size","mfg size"], False),
        "Qty per Pkg/Inner Pack": (["qty per pkg","qty per inner","qty per pkg/inner","qty per pkg/inner pack","qty per pack","qtyperpkg"], True),
        "Pkg Count": (["pkg count","package count","pkg_count","pkgcount","inner pkg count","inner pkg","inner_pack","innerpack","inner pack","pkg"], True),
        "Gross": (["gross","gross weight","grossweight","gross_weight"], False)
    }

    out = pd.DataFrame(index=data.index)
    mapping_log = {}

    for outcol, (cands, prefer_int) in mapping_candidates.items():
        if outcol == "Pkg Count":
            out[outcol] = np.nan
            mapping_log[outcol] = None
            continue
        found = find_col_precise(cols, cands, data, prefer_integer=prefer_int)
        mapping_log[outcol] = found
        out[outcol] = data[found].astype(str).replace(r'^\s*$', np.nan, regex=True) if (found and found in data.columns) else np.nan

    pkg_count_col = None

    for c in cols:
        lc = str(c).strip().lower()
        if re.search(r'\b(inner\s*pkg\s*count|inner\s*pkg|innerpack|inner\s*pack|pkg\s*count|package\s*count|pkgcount)\b', lc):
            pkg_count_col = c
            if debug: run_logs.append(f"[PKG] exact header match -> '{c}'")
            break

    if pkg_count_col is None and mapping_log.get("Qty per Pkg/Inner Pack"):
        try:
            qty_col = mapping_log["Qty per Pkg/Inner Pack"]
            qty_idx = cols.index(qty_col)
            for offset in range(1,5):
                idx = qty_idx + offset
                if idx < len(cols):
                    cand = cols[idx]
                    lc = str(cand).lower()
                    if "pkg" in lc or "inner" in lc or "count" in lc:
                        pkg_count_col = cand
                        if debug: run_logs.append(f"[PKG] relative-position pick -> '{cand}' (offset {offset} from qty col)")
                        break
        except Exception:
            pass

    if pkg_count_col is None:
        for c in cols:
            lc = str(c).lower()
            if (("pkg" in lc) and ("count" in lc)) or (("inner" in lc) and ("pkg" in lc)):
                pkg_count_col = c
                if debug: run_logs.append(f"[PKG] substring fallback -> '{c}'")
                break

    if pkg_count_col is None:
        candidates = []
        for c in cols:
            if c == mapping_log.get("Qty per Pkg/Inner Pack"):
                continue
            ser = data[c].dropna().astype(str).str.strip()
            if ser.empty:
                continue
            nums = pd.to_numeric(ser.str.replace(r'[^\d\-\.]', '', regex=True), errors='coerce')
            if nums.dropna().empty:
                continue
            int_prop = nums.dropna().apply(float.is_integer).mean() if len(nums.dropna())>0 else 0
            prop_ones = (nums==1).mean() if len(nums.dropna())>0 else 0
            score = int_prop*0.6 + prop_ones*0.4
            candidates.append((c, score, prop_ones, int_prop))
        if candidates:
            candidates.sort(key=lambda x: x[1], reverse=True)
            best, best_score, best_prop_one, best_int_prop = candidates[0]
            if best_score > 0.35 and best_prop_one > 0.25:
                pkg_count_col = best
                if debug: run_logs.append(f"[PKG] numeric-heuristic pick -> '{best}' score={best_score:.2f} prop_one={best_prop_one:.2f}")

    if pkg_count_col is None:
        pkg_count_col = find_col_precise(cols, mapping_candidates["Pkg Count"][0], data, prefer_integer=True)
        if pkg_count_col and debug: run_logs.append(f"[PKG] final fallback find_col_precise -> '{pkg_count_col}'")

    mapping_log["Pkg Count"] = pkg_count_col

    if pkg_count_col and pkg_count_col in data.columns:
        raw_pkg = data[pkg_count_col].astype(str).str.strip()
        out["_original_pkg_empty"] = raw_pkg.replace(r'^\s*$', np.nan, regex=True).isna()
        out["Pkg Count"] = pd.to_numeric(raw_pkg.replace(r'^\s*$', np.nan, regex=True), errors="coerce").astype('Int64')
    else:
        out["_original_pkg_empty"] = True
        out["Pkg Count"] = pd.Series([pd.NA]*len(data), index=data.index).astype('Int64')

    try:
        if "Qty per Pkg/Inner Pack" in out.columns and mapping_log.get("Qty per Pkg/Inner Pack"):
            qty_series = pd.to_numeric(out["Qty per Pkg/Inner Pack"], errors="coerce")
            pkg_series = out["Pkg Count"].astype('float')
            if len(pkg_series.dropna())>0 and len(qty_series.dropna())>0:
                overlap = pkg_series.notna() & qty_series.notna()
                if overlap.sum() > 0:
                    same_prop = (pkg_series[overlap] == qty_series[overlap]).mean()
                    if debug: run_logs.append(f"[PKG] same_prop with qty-per-pkg = {same_prop:.2f}")
                    if same_prop > 0.6:
                        alt_candidates = []
                        for c in cols:
                            if c == pkg_count_col or c == mapping_log.get("Qty per Pkg/Inner Pack"):
                                continue
                            lc = str(c).lower()
                            if "pkg" in lc or "inner" in lc or "count" in lc:
                                ser = pd.to_numeric(data[c].astype(str).str.strip().replace(r'^\s*$', np.nan, regex=True), errors='coerce')
                                if ser.dropna().empty:
                                    continue
                                overlap2 = ser.notna() & qty_series.notna()
                                same_prop2 = (ser[overlap2] == qty_series[overlap2]).mean() if overlap2.sum()>0 else 0
                                prop_ones = (ser==1).mean()
                                alt_candidates.append((c, same_prop2, prop_ones))
                        if alt_candidates:
                            alt_candidates.sort(key=lambda x: (x[1], -x[2]))
                            alt = alt_candidates[0][0]
                            if debug: run_logs.append(f"[PKG] switching to alternative candidate '{alt}' (same_prop was too high)")
                            raw_pkg = data[alt].astype(str).str.strip()
                            out["_original_pkg_empty"] = raw_pkg.replace(r'^\s*$', np.nan, regex=True).isna()
                            out["Pkg Count"] = pd.to_numeric(raw_pkg.replace(r'^\s*$', np.nan, regex=True), errors='coerce').astype('Int64')
                            mapping_log["Pkg Count (switched)"] = alt
    except Exception as e:
        if debug: run_logs.append(f"[PKG] fallback switching error: {e}")

    if out["PO #"].isna().all():
        for c in cols:
            if data[c].dropna().astype(str).str.match(r'^\d{6,}$').any():
                out["PO #"] = data[c]
                mapping_log["PO # (heuristic)"] = c
                break

    for c in out.columns:
        if c not in ["Qty per Pkg/Inner Pack", "Pkg Count", "_original_pkg_empty"]:
            out[c] = out[c].astype(str).str.strip().replace('nan', np.nan)

    out["Qty per Pkg/Inner Pack"] = pd.to_numeric(out["Qty per Pkg/Inner Pack"], errors="coerce")
    out["Gross"] = out["Gross"].astype(str).str.strip().replace('nan', np.nan)

    if out["Range"].astype(str).str.contains(r'\d+', na=False).sum() == 0:
        if debug: run_logs.append(f"[WARN] {filename}: no valid Range found. skip.")
        return pd.DataFrame()

    if debug:
        run_logs.append(f"\n[DEBUG] File: {filename}")
        run_logs.append("  Mapping chosen:")
        for k,v in mapping_log.items():
            sample = []
            if v and v in data.columns:
                sample = data[v].dropna().astype(str).unique()[:8].tolist()
            run_logs.append(f"   - {k:25} -> {v} | sample: {sample}")
        run_logs.append("-" * 60)

    return out.reset_index(drop=True)


# --- grouping helper ---
def group_packing_rules(df):
    df = df.copy()
    df["Packing Rule No."] = df["Packing Rule No."].replace('', np.nan)
    valid_rule_mask = df["Packing Rule No."].astype(str).str.match(r'^\d{4}$', na=False)
    df["_ffill_key"] = np.where(valid_rule_mask, df["Packing Rule No."], np.nan)
    df["_ffill_key"] = df["_ffill_key"].ffill()
    detail_row_mask = df.get("_original_pkg_empty", False) == True
    df["Packing Rule No."] = df["_ffill_key"]
    df.loc[detail_row_mask, "Pkg Count"] = pd.NA
    df.loc[detail_row_mask, "Gross Weight"] = np.nan
    df = df.drop(columns=["_ffill_key"])
    return df


# --- add_separator_rows (consecutive-run insertion) ---
def add_separator_rows(df):
    df = df.copy().reset_index(drop=True)
    pr = df['Packing Rule No.'].astype(str).fillna('').str.strip()
    run_id = pr.ne(pr.shift(fill_value='')).cumsum()
    df['_run_id'] = run_id
    df['_run_size'] = df.groupby('_run_id')['Packing Rule No.'].transform('size')
    df['_is_first_of_run'] = pr.ne(pr.shift(fill_value=''))
    df['_insert_before'] = df['_is_first_of_run'] & (df['_run_size'] > 1) & (df.index != 0)
    insert_indices = df[df['_insert_before']].index.tolist()
    helper_cols = ['_run_id','_run_size','_is_first_of_run','_insert_before']
    base_df = df.drop(columns=helper_cols)
    if not insert_indices:
        return base_df
    result_parts = []
    last_idx = 0
    for insert_idx in insert_indices:
        result_parts.append(base_df.iloc[last_idx:insert_idx])
        empty_row = pd.DataFrame([{col: np.nan for col in base_df.columns}])
        result_parts.append(empty_row)
        last_idx = insert_idx
    result_parts.append(base_df.iloc[last_idx:])
    result = pd.concat(result_parts, ignore_index=True)
    return result


# --- apply_multirow_pack_rules (the user's new function) ---
def apply_multirow_pack_rules(df, debug=False):
    df = df.copy()
    for c in ["PO No.","Packing Rule No.","Pkg Count","Gross Weight"]:
        if c not in df.columns:
            df[c] = pd.NA
    df = df.reset_index(drop=True)
    grp_keys = df.groupby(["PO No.","Packing Rule No."]).size().rename("grp_size")
    df = df.merge(grp_keys.reset_index(), on=["PO No.","Packing Rule No."], how="left")
    df["_is_multi_row_group"] = df["grp_size"].fillna(0).astype(int) > 1
    for _, group in df.groupby(["PO No.","Packing Rule No."], sort=False):
        if len(group) > 1:
            idxs = group.index.tolist()
            for i in idxs[1:]:
                df.at[i, "Pkg Count"] = pd.NA
                df.at[i, "Gross Weight"] = np.nan
    df['_group_first'] = False
    for (po, pr), group in df.groupby(["PO No.","Packing Rule No."], sort=False):
        if len(group) > 1:
            first_idx = group.index[0]
            df.at[first_idx, '_group_first'] = True
    insert_indices = df[df['_group_first'] == True].index.tolist()
    df = df.drop(columns=["grp_size", "_is_multi_row_group", "_group_first"])
    if insert_indices:
        result_parts = []
        last_idx = 0
        for insert_idx in insert_indices:
            result_parts.append(df.iloc[last_idx:insert_idx])
            empty_row = pd.DataFrame([{col: np.nan for col in df.columns}])
            result_parts.append(empty_row)
            last_idx = insert_idx
        result_parts.append(df.iloc[last_idx:])
        df = pd.concat(result_parts, ignore_index=True)
    if debug:
        run_logs.append(f"[MULTIROW] Applied multi-row rules. Inserted {len(insert_indices)} separator rows before multi-row groups")
    return df


# --- function to run full pipeline given uploaded files ---
def run_pipeline(packing_files, lookup_file, debug=False, normalize_pr_strip=True, group_ffill=True):
    run_logs.clear()
    results = []

    # extract each packing file
    for up in packing_files:
        b = up.read()
        df_ex = extract_package_detail_bytes(b, filename=up.name, debug=debug)
        if not df_ex.empty:
            results.append(df_ex)
            if debug: run_logs.append(f"  -> {up.name}: extracted {len(df_ex)} rows")
        else:
            if debug: run_logs.append(f"  -> {up.name}: no data extracted")

    if not results:
        raise ValueError("No objects to concatenate — tidak ada data yang berhasil diekstrak. Periksa file / format.")

    df_final = pd.concat(results, ignore_index=True)

    # Normalize columns (rename map)
    rename_map = {
        "PO #": "PO No.",
        "Range": "Packing Rule No.",
        "Buyer Item #": "Style",
        "Manufacturing Size": "Size",
        "Qty per Pkg/Inner Pack": "Qty per Pkg/Inner Pack",
        "Pkg Count": "Pkg Count",
        "Gross": "Gross Weight",
        "_original_pkg_empty": "_original_pkg_empty"
    }
    df_norm = df_final.rename(columns=rename_map)

    for c in ["PO No.", "Packing Rule No.", "Style", "Size", "Qty per Pkg/Inner Pack", "Pkg Count", "Gross Weight", "_original_pkg_empty"]:
        if c not in df_norm.columns:
            if c == "_original_pkg_empty":
                df_norm[c] = False
            else:
                df_norm[c] = None

    if "Source File" in df_norm.columns:
        df_norm = df_norm.drop(columns=["Source File"])

    if "Color" not in df_norm.columns: df_norm["Color"] = None
    if "Item No." not in df_norm.columns: df_norm["Item No."] = None

    cols_order = ["PO No.","Packing Rule No.","Style","Color","Size","Item No.","Qty per Pkg/Inner Pack","Pkg Count","Gross Weight","_original_pkg_empty"]
    df_norm = df_norm[[c for c in cols_order if c in df_norm.columns]]
    df_norm["Gross Weight"] = df_norm["Gross Weight"].astype(str).str.strip().replace('nan', np.nan)

    # Group packing rules (ffill) if requested
    if group_ffill:
        df_norm = group_packing_rules(df_norm)

    # Read lookup
    # Read lookup robustly (support single-sheet and multi-sheet Excel files)
lookup_df = None
if lookup_file is not None:
    _raw = safe_read_excel_bytes(io.BytesIO(lookup_file.read()), sheet_name=None, dtype=str)
    if _raw is None:
        lookup_df = None
    elif isinstance(_raw, dict):
        # try to pick a sensible sheet by common names, otherwise take the first sheet
        for candidate in ("FlexView","FlexView_PGDHDRE","Sheet1","Sheet 1","Lookup","Data","Sheet"):
            if candidate in _raw:
                lookup_df = _raw[candidate]
                break
        if lookup_df is None:
            # fallback: take the first sheet
            lookup_df = list(_raw.values())[0]
    else:
        lookup_df = _raw

if lookup_df is None:
    raise ValueError("Lookup file tidak bisa dibaca atau belum diupload.")

    required_lookup_cols = ["Order #","Material Color Description","Manufacturing Size","UPC/EAN (GTIN)","Country/Region"]
    missing_cols = [c for c in required_lookup_cols if c not in lookup_df.columns]
    if missing_cols:
        raise ValueError(f"Lookup file missing required columns: {missing_cols}")

    for c in required_lookup_cols:
        lookup_df[c] = lookup_df[c].fillna('')

    lookup_df["__Order_norm"] = lookup_df["Order #"].astype(str).str.strip()
    lookup_df["__Size_norm"] = lookup_df["Manufacturing Size"].astype(str).str.strip()
    lookup_df = lookup_df[(lookup_df["__Order_norm"] != '') & (lookup_df["__Order_norm"] != 'nan') & (lookup_df["__Size_norm"] != '') & (lookup_df["__Size_norm"] != 'nan')].copy()
    lookup_df = lookup_df.drop_duplicates(subset=["__Order_norm","__Size_norm"], keep="first")

    # Prepare df for strict merge
    df = df_norm.copy()
    if normalize_pr_strip:
        df["Packing Rule No."] = df["Packing Rule No."].astype(str).str.strip()
    df["__PO_norm"] = df["PO No."].astype(str).str.strip()
    df["__Size_norm"] = df["Size"].astype(str).str.strip()

    df_valid = df[(df["__PO_norm"] != '') & (df["__PO_norm"] != 'nan') & (df["__Size_norm"] != '') & (df["__Size_norm"] != 'nan')].copy()
    df_invalid = df[(df["__PO_norm"] == '') | (df["__PO_norm"] == 'nan') | (df["__Size_norm"] == '') | (df["__Size_norm"] == 'nan')].copy()

    merged = df_valid.merge(
        lookup_df[["__Order_norm","__Size_norm","Material Color Description","UPC/EAN (GTIN)","Country/Region"]],
        left_on=["__PO_norm","__Size_norm"],
        right_on=["__Order_norm","__Size_norm"],
        how="left",
        indicator=True
    )

    matched_count = (merged["_merge"] == "both").sum()
    unmatched_count = (merged["_merge"] == "left_only").sum()
    if debug:
        run_logs.append(f"[MERGE] Matched: {matched_count} rows")
        run_logs.append(f"[MERGE] Unmatched: {unmatched_count} rows")

    merged = merged.drop(columns=["_merge"])

    if len(df_invalid) > 0:
        df_invalid["Material Color Description"] = None
        df_invalid["UPC/EAN (GTIN)"] = None
        df_invalid["Country/Region"] = None
        merged = pd.concat([merged, df_invalid], ignore_index=True)

    merged["Color"] = merged["Material Color Description"].where(merged["Material Color Description"].notna() & (merged["Material Color Description"] != ''), None)
    merged["Item No."] = merged["UPC/EAN (GTIN)"].where(merged["UPC/EAN (GTIN)"].notna() & (merged["UPC/EAN (GTIN)" ] != ''), None)

    final_cols = ["PO No.","Packing Rule No.","Style","Color","Size","Item No.","Qty per Pkg/Inner Pack","Pkg Count","Gross Weight","Country/Region","_original_pkg_empty"]
    final_cols = [c for c in final_cols if c in merged.columns]
    df_final_with_lookup = merged[final_cols].copy()
    df_final_with_lookup["Gross Weight"] = df_final_with_lookup["Gross Weight"].astype(str).str.strip().replace('nan', np.nan)

    if "_original_pkg_empty" in df_final_with_lookup.columns:
        df_final_with_lookup = df_final_with_lookup.drop(columns=["_original_pkg_empty"])

    # Apply multi-row rules (user-supplied function)
    df_final_with_lookup = apply_multirow_pack_rules(df_final_with_lookup, debug=debug)
    run_logs.append(f"After applying multi-row pack rules: {len(df_final_with_lookup)} rows")

    # Also add separators for consecutive single-row runs (if any left)
    df_final_with_lookup = add_separator_rows(df_final_with_lookup)
    run_logs.append(f"After adding separator rows: {len(df_final_with_lookup)} rows")

    # Build unmatched
    unmatched = df_final_with_lookup[(df_final_with_lookup["Color"].isna() | df_final_with_lookup["Item No."].isna())].copy()

    # Export combined & unmatched to in-memory Excel
    out_combined = io.BytesIO()
    with pd.ExcelWriter(out_combined, engine='openpyxl') as writer:
        df_final_with_lookup.to_excel(writer, index=False, sheet_name='Rekap')
        unmatched.to_excel(writer, index=False, sheet_name='Unmatched')
    out_combined.seek(0)

    # Per-PO exports into zipped archive
    po_country_map = lookup_df.set_index("__Order_norm")["Country/Region"].to_dict()
    unique_pos = df_final_with_lookup["PO No."].dropna().unique()

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for po in unique_pos:
            df_po = df_final_with_lookup[df_final_with_lookup["PO No."] == po].copy()
            if "Country/Region" in df_po.columns:
                df_po = df_po.drop(columns=["Country/Region"]) 
            country = po_country_map.get(str(po), "Unknown")
            country = str(country).strip() if pd.notna(country) else "Unknown"
            po_str = str(po)
            country_str = country.replace('/', '-')
            file_name = f"Po import table_ {po_str} {country_str}.xlsx"
            # write excel bytes for this po
            tmp = io.BytesIO()
            with pd.ExcelWriter(tmp, engine='openpyxl') as writer:
                df_po.to_excel(writer, index=False, sheet_name='Rekap')
            tmp.seek(0)
            zf.writestr(file_name, tmp.read())
        # also include combined excel
        zf.writestr(OUTPUT_FILE, out_combined.getvalue())
    zip_buffer.seek(0)

    return df_final_with_lookup, unmatched, out_combined.getvalue(), zip_buffer.getvalue(), run_logs


# ---------------- Run when user clicks ----------------
if run_button:
    if not uploaded_packs:
        st.error("Please upload at least one packing-plan file.")
    elif not uploaded_lookup:
        st.error("Please upload the lookup file.")
    else:
        try:
            with st.spinner("Running pipeline..."):
                df_out, unmatched_out, combined_bytes, zip_bytes, logs = run_pipeline(uploaded_packs, uploaded_lookup, debug=debug, normalize_pr_strip=normalize_pr_strip, group_ffill=auto_group_ffill)

            st.success("Pipeline finished")

            # Show logs
            if debug:
                st.subheader("Run logs")
                for r in logs:
                    st.text(r)

            # Show preview
            st.subheader("Preview — first 200 rows")
            st.dataframe(df_out.head(200))

            st.subheader("Unmatched sample")
            st.dataframe(unmatched_out.head(200))

            # Download buttons
            st.download_button("Download combined Excel (Rekap + Unmatched)", data=combined_bytes, file_name=OUTPUT_FILE, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            st.download_button("Download per-PO ZIP (includes combined)", data=zip_bytes, file_name=os.path.basename('all_outputs.zip'), mime='application/zip')

            # Also allow previewing run-time detection (run_id/run_size) for debugging
            if debug:
                st.subheader("Run detection preview (shows consecutive run_id & run_size)")
                tmp = df_out.copy().reset_index(drop=True)
                tmp['__pr_norm'] = tmp['Packing Rule No.'].astype(str).fillna('').str.strip()
                tmp['__run_id'] = tmp['__pr_norm'].ne(tmp['__pr_norm'].shift(fill_value='')).cumsum()
                tmp['__run_size'] = tmp.groupby('__run_id')['__pr_norm'].transform('size')
                st.dataframe(tmp[['PO No.','Packing Rule No.','Style','Size','Item No.','Pkg Count','Gross Weight','__run_id','__run_size']].head(300))

        except Exception as e:
            st.error(f"Pipeline failed: {e}")
            st.exception(e)
