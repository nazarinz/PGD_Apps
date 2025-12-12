import streamlit as st
import pandas as pd
import numpy as np
import os
import re
import shutil
from io import BytesIO
import zipfile
from datetime import datetime

st.set_page_config(
    page_title="Packing Plan Processor",
    page_icon="📦",
    layout="wide"
)

st.title("📦 Packing Plan Processor")
st.markdown("Process packing plan files and merge with lookup data")

# Sidebar configuration
with st.sidebar:
    st.header("⚙️ Settings")
    debug_mode = st.checkbox("Debug Mode", value=True)
    st.markdown("---")
    st.info("Upload your packing plan files and lookup file to begin processing")

# ============= HELPER FUNCTIONS =============

def norm_size_preserve_dash(x):
    """Trim spaces only — do not drop trailing '-' or change format."""
    if pd.isna(x):
        return ""
    return str(x).strip()

def find_header_row(df, keywords):
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

def extract_package_detail(file_bytes, filename, debug=False):
    try:
        x = pd.read_excel(BytesIO(file_bytes), sheet_name="Detail", header=None, dtype=str)
    except Exception as e:
        if debug:
            st.warning(f"❌ Cannot open sheet 'Detail' in {filename}: {e}")
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
                    if debug:
                        st.info(f"✓ {filename}: header at row {cand}")
                    break

    if header_row is None:
        header_keywords = ["range","serial","buyer","po","pkg","gross","qty per pkg","manufacturing size","buyer item","size"]
        header_row = find_header_row(df, header_keywords)

    if header_row is None:
        if debug:
            st.warning(f"⚠️ {filename}: header not found. Skipping.")
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

    # Pkg Count detection logic - IMPROVED to handle merged cells
    pkg_count_col = None
    
    # Step 1: Look for exact "Pkg Count" column name (case-insensitive, handle merged cells)
    for c in cols:
        lc = str(c).strip().lower()
        # Exact match for "pkg count" variations
        if re.match(r'^(pkg\s*count|package\s*count|pkgcount)

    for c in out.columns:
        if c not in ["Qty per Pkg/Inner Pack", "Pkg Count", "_original_pkg_empty"]:
            out[c] = out[c].astype(str).str.strip().replace('nan', np.nan)

    out["Qty per Pkg/Inner Pack"] = pd.to_numeric(out["Qty per Pkg/Inner Pack"], errors="coerce")
    out["Gross"] = out["Gross"].astype(str).str.strip().replace('nan', np.nan)

    return out.reset_index(drop=True)

def group_packing_rules(df):
    """Group rows without valid Packing Rule No. with the last valid Packing Rule No."""
    df = df.copy()
    df["Packing Rule No."] = df["Packing Rule No."].replace('', np.nan)

    valid_rule_mask = df["Packing Rule No."].astype(str).str.match(r'^\d{4}$', na=False)
    df["_ffill_key"] = np.where(valid_rule_mask, df["Packing Rule No."], np.nan)
    df["_ffill_key"] = df["_ffill_key"].ffill()

    detail_row_mask = df["_original_pkg_empty"] == True
    df["Packing Rule No."] = df["_ffill_key"]

    df.loc[detail_row_mask, "Pkg Count"] = pd.NA
    df.loc[detail_row_mask, "Gross Weight"] = np.nan

    df = df.drop(columns=["_ffill_key"])
    return df

def apply_multirow_pack_rules(df, debug=False):
    """Apply multi-row pack detection rules."""
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

    return df

# ============= STREAMLIT UI =============

col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 Packing Plan Files")
    packing_files = st.file_uploader(
        "Upload Packing Plan Files (.xlsx, .xls)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="packing"
    )
    if packing_files:
        st.success(f"✓ {len(packing_files)} file(s) uploaded")

with col2:
    st.subheader("📋 Lookup File")
    lookup_file = st.file_uploader(
        "Upload Lookup File (.xlsx)",
        type=["xlsx"],
        key="lookup"
    )
    if lookup_file:
        st.success(f"✓ {lookup_file.name} uploaded")

st.markdown("---")

if st.button("🚀 Process Files", type="primary", disabled=not (packing_files and lookup_file)):
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # Step 1: Extract packing files
        status_text.text("📦 Extracting packing plan files...")
        progress_bar.progress(10)
        
        results = []
        for idx, file in enumerate(packing_files):
            file_bytes = file.read()
            df_ex = extract_package_detail(file_bytes, file.name, debug=debug_mode)
            if not df_ex.empty:
                results.append(df_ex)
                if debug_mode:
                    st.info(f"✓ {file.name}: extracted {len(df_ex)} rows")
            progress_bar.progress(10 + (idx + 1) * 20 // len(packing_files))
        
        if not results:
            st.error("❌ No data extracted from packing files. Please check file format.")
            st.stop()
        
        df_final = pd.concat(results, ignore_index=True)
        
        # Step 2: Normalize columns
        status_text.text("🔄 Normalizing data...")
        progress_bar.progress(35)
        
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
        
        if "Color" not in df_norm.columns: df_norm["Color"] = None
        if "Item No." not in df_norm.columns: df_norm["Item No."] = None
        
        cols_order = ["PO No.","Packing Rule No.","Style","Color","Size","Item No.","Qty per Pkg/Inner Pack","Pkg Count","Gross Weight","_original_pkg_empty"]
        df_norm = df_norm[[c for c in cols_order if c in df_norm.columns]]
        
        # Step 3: Group packing rules
        status_text.text("📋 Grouping packing rules...")
        progress_bar.progress(45)
        df_norm = group_packing_rules(df_norm)
        
        # Step 4: Read lookup file
        status_text.text("📖 Reading lookup file...")
        progress_bar.progress(55)
        
        lookup_df = pd.read_excel(lookup_file, dtype=str)
        
        required_lookup_cols = ["Order #","Material Color Description","Manufacturing Size","UPC/EAN (GTIN)","Country/Region"]
        missing_cols = [c for c in required_lookup_cols if c not in lookup_df.columns]
        if missing_cols:
            st.error(f"❌ Lookup file missing required columns: {missing_cols}")
            st.stop()
        
        for c in required_lookup_cols:
            lookup_df[c] = lookup_df[c].fillna('')
        
        lookup_df["__Order_norm"] = lookup_df["Order #"].astype(str).str.strip()
        lookup_df["__Size_norm"] = lookup_df["Manufacturing Size"].astype(str).str.strip()
        
        lookup_df = lookup_df[
            (lookup_df["__Order_norm"] != '') &
            (lookup_df["__Order_norm"] != 'nan') &
            (lookup_df["__Size_norm"] != '') &
            (lookup_df["__Size_norm"] != 'nan')
        ].copy()
        
        lookup_df = lookup_df.drop_duplicates(subset=["__Order_norm","__Size_norm"], keep="first")
        
        # Step 5: Merge with lookup
        status_text.text("🔗 Merging with lookup data...")
        progress_bar.progress(70)
        
        df = df_norm.copy()
        df["__PO_norm"] = df["PO No."].astype(str).str.strip()
        df["__Size_norm"] = df["Size"].astype(str).str.strip()
        
        df_valid = df[
            (df["__PO_norm"] != '') &
            (df["__PO_norm"] != 'nan') &
            (df["__Size_norm"] != '') &
            (df["__Size_norm"] != 'nan')
        ].copy()
        
        df_invalid = df[
            (df["__PO_norm"] == '') |
            (df["__PO_norm"] == 'nan') |
            (df["__Size_norm"] == '') |
            (df["__Size_norm"] == 'nan')
        ].copy()
        
        merged = df_valid.merge(
            lookup_df[["__Order_norm","__Size_norm","Material Color Description","UPC/EAN (GTIN)","Country/Region"]],
            left_on=["__PO_norm","__Size_norm"],
            right_on=["__Order_norm","__Size_norm"],
            how="left",
            indicator=True
        )
        
        matched_count = (merged["_merge"] == "both").sum()
        unmatched_count = (merged["_merge"] == "left_only").sum()
        
        merged = merged.drop(columns=["_merge"])
        
        if len(df_invalid) > 0:
            df_invalid["Material Color Description"] = None
            df_invalid["UPC/EAN (GTIN)"] = None
            df_invalid["Country/Region"] = None
            merged = pd.concat([merged, df_invalid], ignore_index=True)
        
        merged["Color"] = merged["Material Color Description"].where(
            merged["Material Color Description"].notna() & (merged["Material Color Description"] != ''),
            None
        )
        merged["Item No."] = merged["UPC/EAN (GTIN)"].where(
            merged["UPC/EAN (GTIN)"].notna() & (merged["UPC/EAN (GTIN)"] != ''),
            None
        )
        
        final_cols = ["PO No.","Packing Rule No.","Style","Color","Size","Item No.","Qty per Pkg/Inner Pack","Pkg Count","Gross Weight","Country/Region","_original_pkg_empty"]
        final_cols = [c for c in final_cols if c in merged.columns]
        df_final_with_lookup = merged[final_cols].copy()
        
        if "_original_pkg_empty" in df_final_with_lookup.columns:
            df_final_with_lookup = df_final_with_lookup.drop(columns=["_original_pkg_empty"])
        
        # Step 6: Apply multi-row rules
        status_text.text("📊 Applying multi-row pack rules...")
        progress_bar.progress(85)
        df_final_with_lookup = apply_multirow_pack_rules(df_final_with_lookup, debug=debug_mode)
        
        # Step 7: Identify unmatched
        unmatched = df_final_with_lookup[
            (df_final_with_lookup["Color"].isna()) | (df_final_with_lookup["Item No."].isna())
        ].copy()
        
        # Step 8: Create downloads
        status_text.text("💾 Preparing downloads...")
        progress_bar.progress(95)
        
        # Main output file
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final_with_lookup.to_excel(writer, index=False, sheet_name='Rekap')
            if not unmatched.empty:
                unmatched.to_excel(writer, index=False, sheet_name='Unmatched')
        output.seek(0)
        
        # Per-PO exports in ZIP
        po_country_map = lookup_df.set_index("__Order_norm")["Country/Region"].to_dict()
        unique_pos = df_final_with_lookup["PO No."].dropna().unique()
        
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for po in unique_pos:
                df_po = df_final_with_lookup[df_final_with_lookup["PO No."] == po].copy()
                if "Country/Region" in df_po.columns:
                    df_po = df_po.drop(columns=["Country/Region"])
                
                country = po_country_map.get(str(po), "Unknown")
                country = str(country).strip() if pd.notna(country) else "Unknown"
                po_str = str(po)
                country_str = country.replace("/", "-")
                file_name = f"Po import table_ {po_str} {country_str}.xlsx"
                
                po_output = BytesIO()
                with pd.ExcelWriter(po_output, engine='openpyxl') as writer:
                    df_po.to_excel(writer, index=False, sheet_name='Rekap')
                po_output.seek(0)
                
                zip_file.writestr(file_name, po_output.read())
        
        zip_buffer.seek(0)
        
        progress_bar.progress(100)
        status_text.text("✅ Processing complete!")
        
        # Display results
        st.success("🎉 Processing completed successfully!")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Rows", len(df_final_with_lookup))
        with col2:
            filled_color = df_final_with_lookup["Color"].notna().sum()
            st.metric("Color Filled", filled_color)
        with col3:
            filled_item = df_final_with_lookup["Item No."].notna().sum()
            st.metric("Item No. Filled", filled_item)
        with col4:
            st.metric("Unmatched", len(unmatched))
        
        st.markdown("---")
        
        # Download buttons
        col1, col2 = st.columns(2)
        
        with col1:
            st.download_button(
                label="📥 Download Complete Report (Excel)",
                data=output,
                file_name=f"rekap_packing_detail_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col2:
            st.download_button(
                label="📦 Download All PO Exports (ZIP)",
                data=zip_buffer,
                file_name=f"po_exports_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip"
            )
        
        # Preview data
        st.markdown("---")
        st.subheader("📊 Data Preview")
        
        tab1, tab2 = st.tabs(["Complete Data", "Unmatched Rows"])
        
        with tab1:
            st.dataframe(df_final_with_lookup.head(100), use_container_width=True)
        
        with tab2:
            if not unmatched.empty:
                st.dataframe(unmatched.head(100), use_container_width=True)
            else:
                st.info("No unmatched rows found!")
        
    except Exception as e:
        st.error(f"❌ Error during processing: {str(e)}")
        if debug_mode:
            st.exception(e)

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p>Packing Plan Processor v1.0 | Built with Streamlit</p>
</div>
""", unsafe_allow_html=True), lc):
            pkg_count_col = c
            if debug:
                st.info(f"✓ Found Pkg Count column (exact match): '{c}'")
            break
    
    # Step 2: If not found, look for column that contains "pkg count" but not "qty per pkg"
    if pkg_count_col is None:
        for c in cols:
            lc = str(c).strip().lower()
            if 'pkg' in lc and 'count' in lc and 'qty' not in lc and 'per' not in lc:
                pkg_count_col = c
                if debug:
                    st.info(f"✓ Found Pkg Count column (keyword match): '{c}'")
                break
    
    # Step 3: Look in relative position AFTER Qty per Pkg column
    if pkg_count_col is None and mapping_log.get("Qty per Pkg/Inner Pack"):
        try:
            qty_col = mapping_log["Qty per Pkg/Inner Pack"]
            qty_idx = cols.index(qty_col)
            # Check next 1-3 columns after Qty per Pkg
            for offset in range(1, 4):
                idx = qty_idx + offset
                if idx < len(cols):
                    cand = cols[idx]
                    lc = str(cand).lower()
                    # Skip if it's clearly another column (Gross, Weight, etc)
                    if any(skip in lc for skip in ['gross', 'weight', 'kg', 'lbs', 'dimension']):
                        continue
                    # Accept if it has pkg/count/inner keywords or is numeric-heavy
                    if ('pkg' in lc or 'count' in lc or 'inner' in lc) or lc.strip() == '':
                        pkg_count_col = cand
                        if debug:
                            st.info(f"✓ Found Pkg Count column (position-based): '{cand}' at offset +{offset}")
                        break
        except Exception:
            pass
    
    # Step 4: Analyze data patterns - find column with integer counts (often 1s)
    if pkg_count_col is None:
        candidates = []
        qty_col = mapping_log.get("Qty per Pkg/Inner Pack")
        
        for c in cols:
            # Skip already mapped columns
            if c == qty_col or c in [mapping_log.get(k) for k in mapping_log if k != "Pkg Count"]:
                continue
            
            lc = str(c).lower()
            # Skip obvious non-pkg-count columns
            if any(skip in lc for skip in ['gross', 'weight', 'kg', 'dimension', 'serial', 'from', 'to', 'style', 'item', 'color', 'po']):
                continue
            
            ser = data[c].dropna().astype(str).str.strip()
            if ser.empty:
                continue
            
            # Convert to numeric
            nums = pd.to_numeric(ser.str.replace(r'[^\d\-\.]', '', regex=True), errors='coerce').dropna()
            if nums.empty:
                continue
            
            # Calculate scoring metrics
            int_prop = nums.apply(lambda x: float(x).is_integer()).mean()
            prop_small = (nums <= 20).mean()  # Pkg counts usually ≤ 20
            prop_positive = (nums > 0).mean()
            variance = nums.var() if len(nums) > 1 else 0
            
            # Higher score = more likely to be pkg count
            score = (int_prop * 0.4) + (prop_small * 0.3) + (prop_positive * 0.2) + (min(variance, 50) / 50 * 0.1)
            
            candidates.append((c, score, int_prop, prop_small, nums.mean()))
        
        if candidates:
            # Sort by score (descending)
            candidates.sort(key=lambda x: x[1], reverse=True)
            best_col, best_score, best_int, best_small, best_mean = candidates[0]
            
            # Accept if score is reasonable
            if best_score > 0.4:
                pkg_count_col = best_col
                if debug:
                    st.info(f"✓ Found Pkg Count column (pattern analysis): '{best_col}' (score={best_score:.2f}, mean={best_mean:.1f})")
    
    # Step 5: Store the pkg count data
    if pkg_count_col and pkg_count_col in data.columns:
        raw_pkg = data[pkg_count_col].astype(str).str.strip()
        out["_original_pkg_empty"] = raw_pkg.replace(r'^\s*

    for c in out.columns:
        if c not in ["Qty per Pkg/Inner Pack", "Pkg Count", "_original_pkg_empty"]:
            out[c] = out[c].astype(str).str.strip().replace('nan', np.nan)

    out["Qty per Pkg/Inner Pack"] = pd.to_numeric(out["Qty per Pkg/Inner Pack"], errors="coerce")
    out["Gross"] = out["Gross"].astype(str).str.strip().replace('nan', np.nan)

    return out.reset_index(drop=True)

def group_packing_rules(df):
    """Group rows without valid Packing Rule No. with the last valid Packing Rule No."""
    df = df.copy()
    df["Packing Rule No."] = df["Packing Rule No."].replace('', np.nan)

    valid_rule_mask = df["Packing Rule No."].astype(str).str.match(r'^\d{4}$', na=False)
    df["_ffill_key"] = np.where(valid_rule_mask, df["Packing Rule No."], np.nan)
    df["_ffill_key"] = df["_ffill_key"].ffill()

    detail_row_mask = df["_original_pkg_empty"] == True
    df["Packing Rule No."] = df["_ffill_key"]

    df.loc[detail_row_mask, "Pkg Count"] = pd.NA
    df.loc[detail_row_mask, "Gross Weight"] = np.nan

    df = df.drop(columns=["_ffill_key"])
    return df

def apply_multirow_pack_rules(df, debug=False):
    """Apply multi-row pack detection rules."""
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

    return df

# ============= STREAMLIT UI =============

col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 Packing Plan Files")
    packing_files = st.file_uploader(
        "Upload Packing Plan Files (.xlsx, .xls)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="packing"
    )
    if packing_files:
        st.success(f"✓ {len(packing_files)} file(s) uploaded")

with col2:
    st.subheader("📋 Lookup File")
    lookup_file = st.file_uploader(
        "Upload Lookup File (.xlsx)",
        type=["xlsx"],
        key="lookup"
    )
    if lookup_file:
        st.success(f"✓ {lookup_file.name} uploaded")

st.markdown("---")

if st.button("🚀 Process Files", type="primary", disabled=not (packing_files and lookup_file)):
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # Step 1: Extract packing files
        status_text.text("📦 Extracting packing plan files...")
        progress_bar.progress(10)
        
        results = []
        for idx, file in enumerate(packing_files):
            file_bytes = file.read()
            df_ex = extract_package_detail(file_bytes, file.name, debug=debug_mode)
            if not df_ex.empty:
                results.append(df_ex)
                if debug_mode:
                    st.info(f"✓ {file.name}: extracted {len(df_ex)} rows")
            progress_bar.progress(10 + (idx + 1) * 20 // len(packing_files))
        
        if not results:
            st.error("❌ No data extracted from packing files. Please check file format.")
            st.stop()
        
        df_final = pd.concat(results, ignore_index=True)
        
        # Step 2: Normalize columns
        status_text.text("🔄 Normalizing data...")
        progress_bar.progress(35)
        
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
        
        if "Color" not in df_norm.columns: df_norm["Color"] = None
        if "Item No." not in df_norm.columns: df_norm["Item No."] = None
        
        cols_order = ["PO No.","Packing Rule No.","Style","Color","Size","Item No.","Qty per Pkg/Inner Pack","Pkg Count","Gross Weight","_original_pkg_empty"]
        df_norm = df_norm[[c for c in cols_order if c in df_norm.columns]]
        
        # Step 3: Group packing rules
        status_text.text("📋 Grouping packing rules...")
        progress_bar.progress(45)
        df_norm = group_packing_rules(df_norm)
        
        # Step 4: Read lookup file
        status_text.text("📖 Reading lookup file...")
        progress_bar.progress(55)
        
        lookup_df = pd.read_excel(lookup_file, dtype=str)
        
        required_lookup_cols = ["Order #","Material Color Description","Manufacturing Size","UPC/EAN (GTIN)","Country/Region"]
        missing_cols = [c for c in required_lookup_cols if c not in lookup_df.columns]
        if missing_cols:
            st.error(f"❌ Lookup file missing required columns: {missing_cols}")
            st.stop()
        
        for c in required_lookup_cols:
            lookup_df[c] = lookup_df[c].fillna('')
        
        lookup_df["__Order_norm"] = lookup_df["Order #"].astype(str).str.strip()
        lookup_df["__Size_norm"] = lookup_df["Manufacturing Size"].astype(str).str.strip()
        
        lookup_df = lookup_df[
            (lookup_df["__Order_norm"] != '') &
            (lookup_df["__Order_norm"] != 'nan') &
            (lookup_df["__Size_norm"] != '') &
            (lookup_df["__Size_norm"] != 'nan')
        ].copy()
        
        lookup_df = lookup_df.drop_duplicates(subset=["__Order_norm","__Size_norm"], keep="first")
        
        # Step 5: Merge with lookup
        status_text.text("🔗 Merging with lookup data...")
        progress_bar.progress(70)
        
        df = df_norm.copy()
        df["__PO_norm"] = df["PO No."].astype(str).str.strip()
        df["__Size_norm"] = df["Size"].astype(str).str.strip()
        
        df_valid = df[
            (df["__PO_norm"] != '') &
            (df["__PO_norm"] != 'nan') &
            (df["__Size_norm"] != '') &
            (df["__Size_norm"] != 'nan')
        ].copy()
        
        df_invalid = df[
            (df["__PO_norm"] == '') |
            (df["__PO_norm"] == 'nan') |
            (df["__Size_norm"] == '') |
            (df["__Size_norm"] == 'nan')
        ].copy()
        
        merged = df_valid.merge(
            lookup_df[["__Order_norm","__Size_norm","Material Color Description","UPC/EAN (GTIN)","Country/Region"]],
            left_on=["__PO_norm","__Size_norm"],
            right_on=["__Order_norm","__Size_norm"],
            how="left",
            indicator=True
        )
        
        matched_count = (merged["_merge"] == "both").sum()
        unmatched_count = (merged["_merge"] == "left_only").sum()
        
        merged = merged.drop(columns=["_merge"])
        
        if len(df_invalid) > 0:
            df_invalid["Material Color Description"] = None
            df_invalid["UPC/EAN (GTIN)"] = None
            df_invalid["Country/Region"] = None
            merged = pd.concat([merged, df_invalid], ignore_index=True)
        
        merged["Color"] = merged["Material Color Description"].where(
            merged["Material Color Description"].notna() & (merged["Material Color Description"] != ''),
            None
        )
        merged["Item No."] = merged["UPC/EAN (GTIN)"].where(
            merged["UPC/EAN (GTIN)"].notna() & (merged["UPC/EAN (GTIN)"] != ''),
            None
        )
        
        final_cols = ["PO No.","Packing Rule No.","Style","Color","Size","Item No.","Qty per Pkg/Inner Pack","Pkg Count","Gross Weight","Country/Region","_original_pkg_empty"]
        final_cols = [c for c in final_cols if c in merged.columns]
        df_final_with_lookup = merged[final_cols].copy()
        
        if "_original_pkg_empty" in df_final_with_lookup.columns:
            df_final_with_lookup = df_final_with_lookup.drop(columns=["_original_pkg_empty"])
        
        # Step 6: Apply multi-row rules
        status_text.text("📊 Applying multi-row pack rules...")
        progress_bar.progress(85)
        df_final_with_lookup = apply_multirow_pack_rules(df_final_with_lookup, debug=debug_mode)
        
        # Step 7: Identify unmatched
        unmatched = df_final_with_lookup[
            (df_final_with_lookup["Color"].isna()) | (df_final_with_lookup["Item No."].isna())
        ].copy()
        
        # Step 8: Create downloads
        status_text.text("💾 Preparing downloads...")
        progress_bar.progress(95)
        
        # Main output file
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final_with_lookup.to_excel(writer, index=False, sheet_name='Rekap')
            if not unmatched.empty:
                unmatched.to_excel(writer, index=False, sheet_name='Unmatched')
        output.seek(0)
        
        # Per-PO exports in ZIP
        po_country_map = lookup_df.set_index("__Order_norm")["Country/Region"].to_dict()
        unique_pos = df_final_with_lookup["PO No."].dropna().unique()
        
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for po in unique_pos:
                df_po = df_final_with_lookup[df_final_with_lookup["PO No."] == po].copy()
                if "Country/Region" in df_po.columns:
                    df_po = df_po.drop(columns=["Country/Region"])
                
                country = po_country_map.get(str(po), "Unknown")
                country = str(country).strip() if pd.notna(country) else "Unknown"
                po_str = str(po)
                country_str = country.replace("/", "-")
                file_name = f"Po import table_ {po_str} {country_str}.xlsx"
                
                po_output = BytesIO()
                with pd.ExcelWriter(po_output, engine='openpyxl') as writer:
                    df_po.to_excel(writer, index=False, sheet_name='Rekap')
                po_output.seek(0)
                
                zip_file.writestr(file_name, po_output.read())
        
        zip_buffer.seek(0)
        
        progress_bar.progress(100)
        status_text.text("✅ Processing complete!")
        
        # Display results
        st.success("🎉 Processing completed successfully!")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Rows", len(df_final_with_lookup))
        with col2:
            filled_color = df_final_with_lookup["Color"].notna().sum()
            st.metric("Color Filled", filled_color)
        with col3:
            filled_item = df_final_with_lookup["Item No."].notna().sum()
            st.metric("Item No. Filled", filled_item)
        with col4:
            st.metric("Unmatched", len(unmatched))
        
        st.markdown("---")
        
        # Download buttons
        col1, col2 = st.columns(2)
        
        with col1:
            st.download_button(
                label="📥 Download Complete Report (Excel)",
                data=output,
                file_name=f"rekap_packing_detail_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col2:
            st.download_button(
                label="📦 Download All PO Exports (ZIP)",
                data=zip_buffer,
                file_name=f"po_exports_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip"
            )
        
        # Preview data
        st.markdown("---")
        st.subheader("📊 Data Preview")
        
        tab1, tab2 = st.tabs(["Complete Data", "Unmatched Rows"])
        
        with tab1:
            st.dataframe(df_final_with_lookup.head(100), use_container_width=True)
        
        with tab2:
            if not unmatched.empty:
                st.dataframe(unmatched.head(100), use_container_width=True)
            else:
                st.info("No unmatched rows found!")
        
    except Exception as e:
        st.error(f"❌ Error during processing: {str(e)}")
        if debug_mode:
            st.exception(e)

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p>Packing Plan Processor v1.0 | Built with Streamlit</p>
</div>
""", unsafe_allow_html=True), np.nan, regex=True).isna()
        out["Pkg Count"] = pd.to_numeric(raw_pkg.replace(r'^\s*

    for c in out.columns:
        if c not in ["Qty per Pkg/Inner Pack", "Pkg Count", "_original_pkg_empty"]:
            out[c] = out[c].astype(str).str.strip().replace('nan', np.nan)

    out["Qty per Pkg/Inner Pack"] = pd.to_numeric(out["Qty per Pkg/Inner Pack"], errors="coerce")
    out["Gross"] = out["Gross"].astype(str).str.strip().replace('nan', np.nan)

    return out.reset_index(drop=True)

def group_packing_rules(df):
    """Group rows without valid Packing Rule No. with the last valid Packing Rule No."""
    df = df.copy()
    df["Packing Rule No."] = df["Packing Rule No."].replace('', np.nan)

    valid_rule_mask = df["Packing Rule No."].astype(str).str.match(r'^\d{4}$', na=False)
    df["_ffill_key"] = np.where(valid_rule_mask, df["Packing Rule No."], np.nan)
    df["_ffill_key"] = df["_ffill_key"].ffill()

    detail_row_mask = df["_original_pkg_empty"] == True
    df["Packing Rule No."] = df["_ffill_key"]

    df.loc[detail_row_mask, "Pkg Count"] = pd.NA
    df.loc[detail_row_mask, "Gross Weight"] = np.nan

    df = df.drop(columns=["_ffill_key"])
    return df

def apply_multirow_pack_rules(df, debug=False):
    """Apply multi-row pack detection rules."""
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

    return df

# ============= STREAMLIT UI =============

col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 Packing Plan Files")
    packing_files = st.file_uploader(
        "Upload Packing Plan Files (.xlsx, .xls)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="packing"
    )
    if packing_files:
        st.success(f"✓ {len(packing_files)} file(s) uploaded")

with col2:
    st.subheader("📋 Lookup File")
    lookup_file = st.file_uploader(
        "Upload Lookup File (.xlsx)",
        type=["xlsx"],
        key="lookup"
    )
    if lookup_file:
        st.success(f"✓ {lookup_file.name} uploaded")

st.markdown("---")

if st.button("🚀 Process Files", type="primary", disabled=not (packing_files and lookup_file)):
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # Step 1: Extract packing files
        status_text.text("📦 Extracting packing plan files...")
        progress_bar.progress(10)
        
        results = []
        for idx, file in enumerate(packing_files):
            file_bytes = file.read()
            df_ex = extract_package_detail(file_bytes, file.name, debug=debug_mode)
            if not df_ex.empty:
                results.append(df_ex)
                if debug_mode:
                    st.info(f"✓ {file.name}: extracted {len(df_ex)} rows")
            progress_bar.progress(10 + (idx + 1) * 20 // len(packing_files))
        
        if not results:
            st.error("❌ No data extracted from packing files. Please check file format.")
            st.stop()
        
        df_final = pd.concat(results, ignore_index=True)
        
        # Step 2: Normalize columns
        status_text.text("🔄 Normalizing data...")
        progress_bar.progress(35)
        
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
        
        if "Color" not in df_norm.columns: df_norm["Color"] = None
        if "Item No." not in df_norm.columns: df_norm["Item No."] = None
        
        cols_order = ["PO No.","Packing Rule No.","Style","Color","Size","Item No.","Qty per Pkg/Inner Pack","Pkg Count","Gross Weight","_original_pkg_empty"]
        df_norm = df_norm[[c for c in cols_order if c in df_norm.columns]]
        
        # Step 3: Group packing rules
        status_text.text("📋 Grouping packing rules...")
        progress_bar.progress(45)
        df_norm = group_packing_rules(df_norm)
        
        # Step 4: Read lookup file
        status_text.text("📖 Reading lookup file...")
        progress_bar.progress(55)
        
        lookup_df = pd.read_excel(lookup_file, dtype=str)
        
        required_lookup_cols = ["Order #","Material Color Description","Manufacturing Size","UPC/EAN (GTIN)","Country/Region"]
        missing_cols = [c for c in required_lookup_cols if c not in lookup_df.columns]
        if missing_cols:
            st.error(f"❌ Lookup file missing required columns: {missing_cols}")
            st.stop()
        
        for c in required_lookup_cols:
            lookup_df[c] = lookup_df[c].fillna('')
        
        lookup_df["__Order_norm"] = lookup_df["Order #"].astype(str).str.strip()
        lookup_df["__Size_norm"] = lookup_df["Manufacturing Size"].astype(str).str.strip()
        
        lookup_df = lookup_df[
            (lookup_df["__Order_norm"] != '') &
            (lookup_df["__Order_norm"] != 'nan') &
            (lookup_df["__Size_norm"] != '') &
            (lookup_df["__Size_norm"] != 'nan')
        ].copy()
        
        lookup_df = lookup_df.drop_duplicates(subset=["__Order_norm","__Size_norm"], keep="first")
        
        # Step 5: Merge with lookup
        status_text.text("🔗 Merging with lookup data...")
        progress_bar.progress(70)
        
        df = df_norm.copy()
        df["__PO_norm"] = df["PO No."].astype(str).str.strip()
        df["__Size_norm"] = df["Size"].astype(str).str.strip()
        
        df_valid = df[
            (df["__PO_norm"] != '') &
            (df["__PO_norm"] != 'nan') &
            (df["__Size_norm"] != '') &
            (df["__Size_norm"] != 'nan')
        ].copy()
        
        df_invalid = df[
            (df["__PO_norm"] == '') |
            (df["__PO_norm"] == 'nan') |
            (df["__Size_norm"] == '') |
            (df["__Size_norm"] == 'nan')
        ].copy()
        
        merged = df_valid.merge(
            lookup_df[["__Order_norm","__Size_norm","Material Color Description","UPC/EAN (GTIN)","Country/Region"]],
            left_on=["__PO_norm","__Size_norm"],
            right_on=["__Order_norm","__Size_norm"],
            how="left",
            indicator=True
        )
        
        matched_count = (merged["_merge"] == "both").sum()
        unmatched_count = (merged["_merge"] == "left_only").sum()
        
        merged = merged.drop(columns=["_merge"])
        
        if len(df_invalid) > 0:
            df_invalid["Material Color Description"] = None
            df_invalid["UPC/EAN (GTIN)"] = None
            df_invalid["Country/Region"] = None
            merged = pd.concat([merged, df_invalid], ignore_index=True)
        
        merged["Color"] = merged["Material Color Description"].where(
            merged["Material Color Description"].notna() & (merged["Material Color Description"] != ''),
            None
        )
        merged["Item No."] = merged["UPC/EAN (GTIN)"].where(
            merged["UPC/EAN (GTIN)"].notna() & (merged["UPC/EAN (GTIN)"] != ''),
            None
        )
        
        final_cols = ["PO No.","Packing Rule No.","Style","Color","Size","Item No.","Qty per Pkg/Inner Pack","Pkg Count","Gross Weight","Country/Region","_original_pkg_empty"]
        final_cols = [c for c in final_cols if c in merged.columns]
        df_final_with_lookup = merged[final_cols].copy()
        
        if "_original_pkg_empty" in df_final_with_lookup.columns:
            df_final_with_lookup = df_final_with_lookup.drop(columns=["_original_pkg_empty"])
        
        # Step 6: Apply multi-row rules
        status_text.text("📊 Applying multi-row pack rules...")
        progress_bar.progress(85)
        df_final_with_lookup = apply_multirow_pack_rules(df_final_with_lookup, debug=debug_mode)
        
        # Step 7: Identify unmatched
        unmatched = df_final_with_lookup[
            (df_final_with_lookup["Color"].isna()) | (df_final_with_lookup["Item No."].isna())
        ].copy()
        
        # Step 8: Create downloads
        status_text.text("💾 Preparing downloads...")
        progress_bar.progress(95)
        
        # Main output file
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final_with_lookup.to_excel(writer, index=False, sheet_name='Rekap')
            if not unmatched.empty:
                unmatched.to_excel(writer, index=False, sheet_name='Unmatched')
        output.seek(0)
        
        # Per-PO exports in ZIP
        po_country_map = lookup_df.set_index("__Order_norm")["Country/Region"].to_dict()
        unique_pos = df_final_with_lookup["PO No."].dropna().unique()
        
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for po in unique_pos:
                df_po = df_final_with_lookup[df_final_with_lookup["PO No."] == po].copy()
                if "Country/Region" in df_po.columns:
                    df_po = df_po.drop(columns=["Country/Region"])
                
                country = po_country_map.get(str(po), "Unknown")
                country = str(country).strip() if pd.notna(country) else "Unknown"
                po_str = str(po)
                country_str = country.replace("/", "-")
                file_name = f"Po import table_ {po_str} {country_str}.xlsx"
                
                po_output = BytesIO()
                with pd.ExcelWriter(po_output, engine='openpyxl') as writer:
                    df_po.to_excel(writer, index=False, sheet_name='Rekap')
                po_output.seek(0)
                
                zip_file.writestr(file_name, po_output.read())
        
        zip_buffer.seek(0)
        
        progress_bar.progress(100)
        status_text.text("✅ Processing complete!")
        
        # Display results
        st.success("🎉 Processing completed successfully!")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Rows", len(df_final_with_lookup))
        with col2:
            filled_color = df_final_with_lookup["Color"].notna().sum()
            st.metric("Color Filled", filled_color)
        with col3:
            filled_item = df_final_with_lookup["Item No."].notna().sum()
            st.metric("Item No. Filled", filled_item)
        with col4:
            st.metric("Unmatched", len(unmatched))
        
        st.markdown("---")
        
        # Download buttons
        col1, col2 = st.columns(2)
        
        with col1:
            st.download_button(
                label="📥 Download Complete Report (Excel)",
                data=output,
                file_name=f"rekap_packing_detail_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col2:
            st.download_button(
                label="📦 Download All PO Exports (ZIP)",
                data=zip_buffer,
                file_name=f"po_exports_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip"
            )
        
        # Preview data
        st.markdown("---")
        st.subheader("📊 Data Preview")
        
        tab1, tab2 = st.tabs(["Complete Data", "Unmatched Rows"])
        
        with tab1:
            st.dataframe(df_final_with_lookup.head(100), use_container_width=True)
        
        with tab2:
            if not unmatched.empty:
                st.dataframe(unmatched.head(100), use_container_width=True)
            else:
                st.info("No unmatched rows found!")
        
    except Exception as e:
        st.error(f"❌ Error during processing: {str(e)}")
        if debug_mode:
            st.exception(e)

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p>Packing Plan Processor v1.0 | Built with Streamlit</p>
</div>
""", unsafe_allow_html=True), np.nan, regex=True), errors="coerce").astype('Int64')
        mapping_log["Pkg Count"] = pkg_count_col
        if debug:
            st.success(f"✓ Pkg Count column mapped: '{pkg_count_col}'")
    else:
        out["_original_pkg_empty"] = True
        out["Pkg Count"] = pd.Series([pd.NA]*len(data), index=data.index).astype('Int64')
        mapping_log["Pkg Count"] = None
        if debug:
            st.warning("⚠️ Pkg Count column not found - will be empty")

    for c in out.columns:
        if c not in ["Qty per Pkg/Inner Pack", "Pkg Count", "_original_pkg_empty"]:
            out[c] = out[c].astype(str).str.strip().replace('nan', np.nan)

    out["Qty per Pkg/Inner Pack"] = pd.to_numeric(out["Qty per Pkg/Inner Pack"], errors="coerce")
    out["Gross"] = out["Gross"].astype(str).str.strip().replace('nan', np.nan)

    return out.reset_index(drop=True)

def group_packing_rules(df):
    """Group rows without valid Packing Rule No. with the last valid Packing Rule No."""
    df = df.copy()
    df["Packing Rule No."] = df["Packing Rule No."].replace('', np.nan)

    valid_rule_mask = df["Packing Rule No."].astype(str).str.match(r'^\d{4}$', na=False)
    df["_ffill_key"] = np.where(valid_rule_mask, df["Packing Rule No."], np.nan)
    df["_ffill_key"] = df["_ffill_key"].ffill()

    detail_row_mask = df["_original_pkg_empty"] == True
    df["Packing Rule No."] = df["_ffill_key"]

    df.loc[detail_row_mask, "Pkg Count"] = pd.NA
    df.loc[detail_row_mask, "Gross Weight"] = np.nan

    df = df.drop(columns=["_ffill_key"])
    return df

def apply_multirow_pack_rules(df, debug=False):
    """Apply multi-row pack detection rules."""
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

    return df

# ============= STREAMLIT UI =============

col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 Packing Plan Files")
    packing_files = st.file_uploader(
        "Upload Packing Plan Files (.xlsx, .xls)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="packing"
    )
    if packing_files:
        st.success(f"✓ {len(packing_files)} file(s) uploaded")

with col2:
    st.subheader("📋 Lookup File")
    lookup_file = st.file_uploader(
        "Upload Lookup File (.xlsx)",
        type=["xlsx"],
        key="lookup"
    )
    if lookup_file:
        st.success(f"✓ {lookup_file.name} uploaded")

st.markdown("---")

if st.button("🚀 Process Files", type="primary", disabled=not (packing_files and lookup_file)):
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # Step 1: Extract packing files
        status_text.text("📦 Extracting packing plan files...")
        progress_bar.progress(10)
        
        results = []
        for idx, file in enumerate(packing_files):
            file_bytes = file.read()
            df_ex = extract_package_detail(file_bytes, file.name, debug=debug_mode)
            if not df_ex.empty:
                results.append(df_ex)
                if debug_mode:
                    st.info(f"✓ {file.name}: extracted {len(df_ex)} rows")
            progress_bar.progress(10 + (idx + 1) * 20 // len(packing_files))
        
        if not results:
            st.error("❌ No data extracted from packing files. Please check file format.")
            st.stop()
        
        df_final = pd.concat(results, ignore_index=True)
        
        # Step 2: Normalize columns
        status_text.text("🔄 Normalizing data...")
        progress_bar.progress(35)
        
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
        
        if "Color" not in df_norm.columns: df_norm["Color"] = None
        if "Item No." not in df_norm.columns: df_norm["Item No."] = None
        
        cols_order = ["PO No.","Packing Rule No.","Style","Color","Size","Item No.","Qty per Pkg/Inner Pack","Pkg Count","Gross Weight","_original_pkg_empty"]
        df_norm = df_norm[[c for c in cols_order if c in df_norm.columns]]
        
        # Step 3: Group packing rules
        status_text.text("📋 Grouping packing rules...")
        progress_bar.progress(45)
        df_norm = group_packing_rules(df_norm)
        
        # Step 4: Read lookup file
        status_text.text("📖 Reading lookup file...")
        progress_bar.progress(55)
        
        lookup_df = pd.read_excel(lookup_file, dtype=str)
        
        required_lookup_cols = ["Order #","Material Color Description","Manufacturing Size","UPC/EAN (GTIN)","Country/Region"]
        missing_cols = [c for c in required_lookup_cols if c not in lookup_df.columns]
        if missing_cols:
            st.error(f"❌ Lookup file missing required columns: {missing_cols}")
            st.stop()
        
        for c in required_lookup_cols:
            lookup_df[c] = lookup_df[c].fillna('')
        
        lookup_df["__Order_norm"] = lookup_df["Order #"].astype(str).str.strip()
        lookup_df["__Size_norm"] = lookup_df["Manufacturing Size"].astype(str).str.strip()
        
        lookup_df = lookup_df[
            (lookup_df["__Order_norm"] != '') &
            (lookup_df["__Order_norm"] != 'nan') &
            (lookup_df["__Size_norm"] != '') &
            (lookup_df["__Size_norm"] != 'nan')
        ].copy()
        
        lookup_df = lookup_df.drop_duplicates(subset=["__Order_norm","__Size_norm"], keep="first")
        
        # Step 5: Merge with lookup
        status_text.text("🔗 Merging with lookup data...")
        progress_bar.progress(70)
        
        df = df_norm.copy()
        df["__PO_norm"] = df["PO No."].astype(str).str.strip()
        df["__Size_norm"] = df["Size"].astype(str).str.strip()
        
        df_valid = df[
            (df["__PO_norm"] != '') &
            (df["__PO_norm"] != 'nan') &
            (df["__Size_norm"] != '') &
            (df["__Size_norm"] != 'nan')
        ].copy()
        
        df_invalid = df[
            (df["__PO_norm"] == '') |
            (df["__PO_norm"] == 'nan') |
            (df["__Size_norm"] == '') |
            (df["__Size_norm"] == 'nan')
        ].copy()
        
        merged = df_valid.merge(
            lookup_df[["__Order_norm","__Size_norm","Material Color Description","UPC/EAN (GTIN)","Country/Region"]],
            left_on=["__PO_norm","__Size_norm"],
            right_on=["__Order_norm","__Size_norm"],
            how="left",
            indicator=True
        )
        
        matched_count = (merged["_merge"] == "both").sum()
        unmatched_count = (merged["_merge"] == "left_only").sum()
        
        merged = merged.drop(columns=["_merge"])
        
        if len(df_invalid) > 0:
            df_invalid["Material Color Description"] = None
            df_invalid["UPC/EAN (GTIN)"] = None
            df_invalid["Country/Region"] = None
            merged = pd.concat([merged, df_invalid], ignore_index=True)
        
        merged["Color"] = merged["Material Color Description"].where(
            merged["Material Color Description"].notna() & (merged["Material Color Description"] != ''),
            None
        )
        merged["Item No."] = merged["UPC/EAN (GTIN)"].where(
            merged["UPC/EAN (GTIN)"].notna() & (merged["UPC/EAN (GTIN)"] != ''),
            None
        )
        
        final_cols = ["PO No.","Packing Rule No.","Style","Color","Size","Item No.","Qty per Pkg/Inner Pack","Pkg Count","Gross Weight","Country/Region","_original_pkg_empty"]
        final_cols = [c for c in final_cols if c in merged.columns]
        df_final_with_lookup = merged[final_cols].copy()
        
        if "_original_pkg_empty" in df_final_with_lookup.columns:
            df_final_with_lookup = df_final_with_lookup.drop(columns=["_original_pkg_empty"])
        
        # Step 6: Apply multi-row rules
        status_text.text("📊 Applying multi-row pack rules...")
        progress_bar.progress(85)
        df_final_with_lookup = apply_multirow_pack_rules(df_final_with_lookup, debug=debug_mode)
        
        # Step 7: Identify unmatched
        unmatched = df_final_with_lookup[
            (df_final_with_lookup["Color"].isna()) | (df_final_with_lookup["Item No."].isna())
        ].copy()
        
        # Step 8: Create downloads
        status_text.text("💾 Preparing downloads...")
        progress_bar.progress(95)
        
        # Main output file
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final_with_lookup.to_excel(writer, index=False, sheet_name='Rekap')
            if not unmatched.empty:
                unmatched.to_excel(writer, index=False, sheet_name='Unmatched')
        output.seek(0)
        
        # Per-PO exports in ZIP
        po_country_map = lookup_df.set_index("__Order_norm")["Country/Region"].to_dict()
        unique_pos = df_final_with_lookup["PO No."].dropna().unique()
        
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for po in unique_pos:
                df_po = df_final_with_lookup[df_final_with_lookup["PO No."] == po].copy()
                if "Country/Region" in df_po.columns:
                    df_po = df_po.drop(columns=["Country/Region"])
                
                country = po_country_map.get(str(po), "Unknown")
                country = str(country).strip() if pd.notna(country) else "Unknown"
                po_str = str(po)
                country_str = country.replace("/", "-")
                file_name = f"Po import table_ {po_str} {country_str}.xlsx"
                
                po_output = BytesIO()
                with pd.ExcelWriter(po_output, engine='openpyxl') as writer:
                    df_po.to_excel(writer, index=False, sheet_name='Rekap')
                po_output.seek(0)
                
                zip_file.writestr(file_name, po_output.read())
        
        zip_buffer.seek(0)
        
        progress_bar.progress(100)
        status_text.text("✅ Processing complete!")
        
        # Display results
        st.success("🎉 Processing completed successfully!")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Rows", len(df_final_with_lookup))
        with col2:
            filled_color = df_final_with_lookup["Color"].notna().sum()
            st.metric("Color Filled", filled_color)
        with col3:
            filled_item = df_final_with_lookup["Item No."].notna().sum()
            st.metric("Item No. Filled", filled_item)
        with col4:
            st.metric("Unmatched", len(unmatched))
        
        st.markdown("---")
        
        # Download buttons
        col1, col2 = st.columns(2)
        
        with col1:
            st.download_button(
                label="📥 Download Complete Report (Excel)",
                data=output,
                file_name=f"rekap_packing_detail_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col2:
            st.download_button(
                label="📦 Download All PO Exports (ZIP)",
                data=zip_buffer,
                file_name=f"po_exports_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip"
            )
        
        # Preview data
        st.markdown("---")
        st.subheader("📊 Data Preview")
        
        tab1, tab2 = st.tabs(["Complete Data", "Unmatched Rows"])
        
        with tab1:
            st.dataframe(df_final_with_lookup.head(100), use_container_width=True)
        
        with tab2:
            if not unmatched.empty:
                st.dataframe(unmatched.head(100), use_container_width=True)
            else:
                st.info("No unmatched rows found!")
        
    except Exception as e:
        st.error(f"❌ Error during processing: {str(e)}")
        if debug_mode:
            st.exception(e)

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p>Packing Plan Processor v1.0 | Built with Streamlit</p>
</div>
""", unsafe_allow_html=True)
