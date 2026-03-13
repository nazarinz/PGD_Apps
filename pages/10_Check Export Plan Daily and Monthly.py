from utils.auth import require_login

require_login()

# pages/2_Check Export Plan Daily and Monthly.py
# PGD Apps — Check Export Plan Daily and Monthly (SO Auto-Detect) + Drill-down Matches (cached)
import io, re
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Set

import pandas as pd
import streamlit as st

st.set_page_config(page_title="PGD Apps — Check Export Plan Daily and Monthly", page_icon="🔎", layout="wide")
st.title("🔎 SO Auto-Detect — Export Plan Daily and Monthly")

# =========================
# Utils: Baca file & deteksi kolom SO
# =========================
def _read_excel_or_ods(file_bytes: bytes, suffix: str) -> Dict[str, pd.DataFrame]:
    b = io.BytesIO(file_bytes)
    sfx = suffix.lower()
    if sfx in (".xlsx", ".xlsm"):
        return pd.read_excel(b, engine="openpyxl", sheet_name=None)
    elif sfx == ".xls":
        try:
            return pd.read_excel(io.BytesIO(file_bytes), engine="xlrd", sheet_name=None)
        except Exception:
            try:
                return pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl", sheet_name=None)
            except Exception as e2:
                raise RuntimeError("Gagal membaca .xls. Pastikan file tidak korup atau misnamed.") from e2
    elif sfx == ".xlsb":
        try:
            return pd.read_excel(b, engine="pyxlsb", sheet_name=None)
        except Exception as e:
            raise RuntimeError("Gagal membaca .xlsb. Install 'pyxlsb' atau konversi ke .xlsx.") from e
    elif sfx == ".ods":
        try:
            return pd.read_excel(b, engine="odf", sheet_name=None)
        except Exception as e:
            raise RuntimeError("Gagal membaca ODS. Install 'odfpy' atau konversi ke .xlsx.") from e
    else:
        raise ValueError("Unsupported Excel format")

def read_any(uploaded_file) -> Dict[str, pd.DataFrame]:
    name = uploaded_file.name
    suf = name.strip().lower()
    if suf.endswith(".csv"):
        try:
            uploaded_file.seek(0)
        except Exception:
            pass
        return {"Sheet1": pd.read_csv(uploaded_file)}
    elif suf.endswith((".xlsx", ".xls", ".ods", ".xlsb", ".xlsm")):
        if hasattr(uploaded_file, "getvalue"):
            raw_bytes = uploaded_file.getvalue()
        else:
            try:
                uploaded_file.seek(0)
            except Exception:
                pass
            raw_bytes = uploaded_file.read()
        ext = "." + name.split(".")[-1].lower()
        return _read_excel_or_ods(raw_bytes, ext)
    else:
        raise ValueError("Format tidak didukung. Gunakan CSV/XLSX/XLS/XLSB/XLSM/ODS.")

CANDIDATES = [
    "sap_odr_no", "sap_odrno", "sap_order_no", "sap_odr_number",
    "sales_order", "salesorder", "so", "so_no", "sono", "so_number",
]

# Kandidat khusus untuk kolom FVB SO di base file
FVB_CANDIDATES = [
    "fvb_so", "fvbso", "fvb so", "fvb_sales_order", "fvb",
]

def norm_col_name(c: str) -> str:
    c = str(c).strip().lower()
    c = re.sub(r"\s+", " ", c)
    c = re.sub(r"[^0-9a-z_ ]+", "", c)
    c = c.replace(" ", "_")
    c = re.sub(r"_+", "_", c)
    return c

def normalize_columns(df: pd.DataFrame):
    mapping = {}
    new_cols = []
    for c in df.columns:
        nc = norm_col_name(c)
        base = nc or "col"
        counter = 1
        while base in mapping:
            counter += 1
            base = f"{nc}_{counter}"
        mapping[base] = str(c)
        new_cols.append(base)
    out = df.copy()
    out.columns = new_cols
    return out, mapping

def detect_so_column(df: pd.DataFrame):
    if df is None or df.empty or (df.dropna(how="all").empty):
        return None, "DataFrame kosong"
    df_n, mapping = normalize_columns(df)
    cols_norm = list(df_n.columns)

    # 1) kandidat eksak
    for cand in CANDIDATES:
        if cand in cols_norm:
            return mapping[cand], f"Exact candidate match: {cand}"
    # 2) heuristik
    for cn in cols_norm:
        if ("sap" in cn) and ("odr" in cn or "order" in cn):
            return mapping[cn], f"Heuristic: contains 'sap' and 'odr/order' → {cn}"
    for cn in cols_norm:
        if ("sales" in cn) and ("order" in cn):
            return mapping[cn], f"Heuristic: contains 'sales' and 'order' → {cn}"
    for cn in cols_norm:
        if cn in {"order","so"}:
            return mapping[cn], f"Fallback exact short name: {cn}"

    # 3) sniff header di beberapa baris awal
    tokens = ["sap odr no", "sap order", "sales order", "so", "so no", "so number"]
    max_scan = min(len(df), 10)
    for r in range(max_scan):
        row_vals = [str(x) for x in df.iloc[r, :].tolist()]
        combined = " | ".join(row_vals).lower()
        if any(tok in combined for tok in tokens):
            tmp = df.iloc[r+1:].copy()
            tmp.columns = [norm_col_name(x) for x in row_vals]
            tmp_map = {norm_col_name(x): str(x) for x in row_vals}
            for cand in CANDIDATES:
                if cand in tmp.columns:
                    return tmp_map[cand], f"Header-sniff row {r}: {cand}"
            for cn in tmp.columns:
                if ("sap" in cn) and ("odr" in cn or "order" in cn):
                    return tmp_map[cn], f"Header-sniff row {r}: sap+odr/order → {cn}"
            for cn in tmp.columns:
                if ("sales" in cn) and ("order" in cn):
                    return tmp_map[cn], f"Header-sniff row {r}: sales+order → {cn}"
            for cn in tmp.columns:
                if cn in {"order","so"}:
                    return tmp_map[cn], f"Header-sniff row {r}: fallback {cn}"
    return None, "Tidak ditemukan kolom yang cocok"


def detect_fvb_so_column(df: pd.DataFrame) -> Tuple[Optional[str], str]:
    """Deteksi kolom FVB SO khusus untuk base file."""
    if df is None or df.empty:
        return None, "DataFrame kosong"
    df_n, mapping = normalize_columns(df)
    cols_norm = list(df_n.columns)

    # 1) Eksak dari daftar kandidat FVB
    for cand in FVB_CANDIDATES:
        nc = norm_col_name(cand)
        if nc in cols_norm:
            return mapping[nc], f"Exact FVB candidate match: {nc}"

    # 2) Heuristik: kolom yang mengandung "fvb"
    for cn in cols_norm:
        if "fvb" in cn:
            return mapping[cn], f"Heuristic: contains 'fvb' → {cn}"

    return None, "Kolom FVB SO tidak ditemukan"


def reheader_with_row(df: pd.DataFrame, header_row_index: int = 2) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    if header_row_index >= len(df):
        return df
    raw_header = [str(x) if x is not None else "" for x in df.iloc[header_row_index].tolist()]
    seen = set()
    uniq_cols = []
    for c in raw_header:
        base = c.strip() or "col"
        key = base
        n = 1
        while key in seen:
            n += 1
            key = f"{base}_{n}"
        uniq_cols.append(key)
        seen.add(key)
    df2 = df.iloc[header_row_index+1:].copy()
    df2.columns = uniq_cols
    df2 = df2.dropna(how="all")
    df2 = df2.dropna(axis=1, how="all")
    return df2

def pick_column(df: pd.DataFrame, target_name: str) -> Optional[str]:
    if df is None or df.empty:
        return None
    for c in df.columns:
        if str(c) == target_name:
            return str(c)
    target_lower = target_name.lower()
    for c in df.columns:
        if str(c).lower() == target_lower:
            return str(c)
    norm_target = norm_col_name(target_name)
    for c in df.columns:
        if norm_col_name(c) == norm_target:
            return str(c)
    return None

def normalize_so_series(s: pd.Series) -> pd.Series:
    s2 = s.astype(str).str.strip()
    s2 = s2.str.replace(r"\.0$", "", regex=True)
    s2 = s2.replace({"nan": pd.NA, "none": pd.NA, "": pd.NA})
    return s2

# =========================
# Struktur hasil (untuk cache)
# =========================
@dataclass
class SOResult:
    base_file_name: str
    base_sheet_name: str
    # --- Kolom SO pertama (SAP_ODR_NO) ---
    base_col: str
    base_reason: str
    # --- Kolom SO kedua (FVB SO) ---
    base_col_fvb: Optional[str]
    base_reason_fvb: str
    ref_count: int
    matches: pd.DataFrame
    not_found: pd.DataFrame
    empty_so: pd.DataFrame
    base_df_out: pd.DataFrame        # punya kolom __SO_norm_SAP__ dan __SO_norm_FVB__
    ref_tables: list
    log_df: pd.DataFrame
    matched_so_list: list            # gabungan SO match dari kedua kolom

# =========================
# Proses utama (sekali klik)
# =========================
def process_files(base_file, ref_files) -> SOResult:
    # Cari sheet base (prefer "Loading Plan")
    base_sheets = read_any(base_file)
    base_sheet_name = None
    for sh in base_sheets.keys():
        if str(sh).strip().lower().replace(" ", "_") == "loading_plan":
            base_sheet_name = sh
            break
    if base_sheet_name is None:
        base_sheet_name = next(iter(base_sheets.keys()))

    if hasattr(base_file, "getvalue"):
        raw_bytes = base_file.getvalue()
    else:
        try:
            base_file.seek(0)
        except Exception:
            pass
        raw_bytes = base_file.read()

    ext = "." + base_file.name.split(".")[-1].lower()
    engine_map = {".ods": "odf", ".xlsx": "openpyxl", ".xlsm": "openpyxl",
                  ".xls": "xlrd", ".xlsb": "pyxlsb"}
    engine = engine_map.get(ext, "openpyxl")

    book_headerless = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=None, engine=engine, header=None)
    _base_df_raw = book_headerless[base_sheet_name]
    base_df = reheader_with_row(_base_df_raw, header_row_index=2)

    # ── Kolom SO #1: SAP_ODR_NO ──────────────────────────────────────────────
    base_col = pick_column(base_df, "SAP_ODR_NO")
    base_reason = "Forced: SAP_ODR_NO | header=baris 3"
    if not base_col:
        base_col, base_reason = detect_so_column(base_df)
        base_reason = f"{base_reason} | header=baris 3"
    if base_col is None:
        base_col = base_df.columns[0]
        base_reason = f"Fallback: pakai kolom pertama ({base_col})"

    # ── Kolom SO #2: FVB SO ───────────────────────────────────────────────────
    base_col_fvb = pick_column(base_df, "FVB SO")
    if not base_col_fvb:
        base_col_fvb, base_reason_fvb = detect_fvb_so_column(base_df)
    else:
        base_reason_fvb = "Forced: FVB SO | header=baris 3"
    if base_col_fvb is None:
        base_reason_fvb = "Kolom FVB SO tidak ditemukan di base file"
    # ─────────────────────────────────────────────────────────────────────────

    base_df_out = base_df.copy()
    base_df_out["__SO_norm_SAP__"] = normalize_so_series(base_df_out[base_col])
    if base_col_fvb and base_col_fvb in base_df_out.columns:
        base_df_out["__SO_norm_FVB__"] = normalize_so_series(base_df_out[base_col_fvb])
    else:
        base_df_out["__SO_norm_FVB__"] = pd.NA

    # Kumpulkan SO referensi
    all_ref_sos: Set[str] = set()
    so_to_files: Dict[str, Set[str]] = {}
    log_rows: List[Tuple[str, str, str]] = []
    ref_tables: List[Dict[str, object]] = []

    if ref_files:
        for f in ref_files:
            sheets = read_any(f)
            for sh, df in sheets.items():
                df2 = df.copy()
                if not df2.empty:
                    df2 = df2.dropna(how="all")
                    df2 = df2.dropna(axis=1, how="all")

                col, reason = detect_so_column(df2)
                log_rows.append((f.name, sh, reason))

                if not col or col not in df2.columns:
                    if col:
                        log_rows.append((f.name, sh, f"Kolom '{col}' terdeteksi tapi hilang setelah cleanup"))
                    continue

                df2["__SO_norm__"] = normalize_so_series(df2[col])
                sos = df2["__SO_norm__"].dropna()
                if sos.empty:
                    log_rows.append((f.name, sh, f"Kolom '{col}' terdeteksi tapi tidak ada nilai SO valid"))
                    continue

                for so in sos.unique().tolist():
                    all_ref_sos.add(so)
                    so_to_files.setdefault(so, set()).add(f.name)

                ref_tables.append({
                    "file": f.name,
                    "sheet": sh,
                    "so_col": col,
                    "df": df2,
                })

    # ── Matching: cek KEDUA kolom SO terhadap referensi ──────────────────────
    out = base_df_out.copy()

    def _found(row):
        sap = row["__SO_norm_SAP__"]
        fvb = row["__SO_norm_FVB__"]
        return (pd.notna(sap) and sap in all_ref_sos) or (pd.notna(fvb) and fvb in all_ref_sos)

    def _source_files(row):
        files = set()
        sap = row["__SO_norm_SAP__"]
        fvb = row["__SO_norm_FVB__"]
        if pd.notna(sap) and sap in so_to_files:
            files |= so_to_files[sap]
        if pd.notna(fvb) and fvb in so_to_files:
            files |= so_to_files[fvb]
        return ", ".join(sorted(files)) if files else ""

    out["Found_in_reference"] = out.apply(_found, axis=1)
    out["Source_Files"]       = out.apply(_source_files, axis=1)
    # ─────────────────────────────────────────────────────────────────────────

    matches   = out[out["Found_in_reference"] == True].copy()
    not_found = out[
        (~out["Found_in_reference"]) &
        (out["__SO_norm_SAP__"].notna() | out["__SO_norm_FVB__"].notna())
    ].copy()
    empty_so = out[
        out["__SO_norm_SAP__"].isna() & out["__SO_norm_FVB__"].isna()
    ].copy()

    # Kumpulkan semua SO yang match dari kedua kolom
    # Filter dengan all_ref_sos agar hanya SO yang benar-benar ada di referensi
    matched_sap = set(matches["__SO_norm_SAP__"].dropna().astype(str).unique())
    matched_fvb = set(matches["__SO_norm_FVB__"].dropna().astype(str).unique())
    matched_so_list = sorted((matched_sap | matched_fvb) & all_ref_sos)

    log_df = (pd.DataFrame(log_rows, columns=["File", "Sheet", "Reason"])
              if log_rows else pd.DataFrame(columns=["File","Sheet","Reason"]))

    return SOResult(
        base_file_name  = base_file.name,
        base_sheet_name = base_sheet_name,
        base_col        = base_col,
        base_reason     = base_reason,
        base_col_fvb    = base_col_fvb,
        base_reason_fvb = base_reason_fvb,
        ref_count       = len(ref_files or []),
        matches         = matches,
        not_found       = not_found,
        empty_so        = empty_so,
        base_df_out     = base_df_out,
        ref_tables      = ref_tables,
        log_df          = log_df,
        matched_so_list = matched_so_list,
    )

# =========================
# UI Upload
# =========================
col1, col2 = st.columns(2)
with col1:
    base_file = st.file_uploader(
        "Daily file (CSV/XLSX/XLS/ODS/XLSB/XLSM)",
        type=["csv","xlsx","xls","ods","xlsb","xlsm"],
        accept_multiple_files=False
    )
with col2:
    ref_files = st.file_uploader(
        "Reference files (bisa banyak)",
        type=["csv","xlsx","xls","ods","xlsb","xlsm"],
        accept_multiple_files=True
    )

colA, colB = st.columns([1,1])
with colA:
    do_process = st.button("▶️ Proses", type="primary")
with colB:
    do_reset = st.button("♻️ Reset hasil")

if do_reset:
    st.session_state.pop("so_detect_result", None)
    st.success("Hasil dipulihkan. Silakan klik Proses lagi setelah pilih file.")

if do_process:
    if not base_file:
        st.error("Unggah dulu **Daily file**.")
        st.stop()
    st.session_state["so_detect_result"] = process_files(base_file, ref_files)

# =========================
# Render dari cache
# =========================
if "so_detect_result" in st.session_state:
    res: SOResult = st.session_state["so_detect_result"]

    st.subheader("📄 Daily file (base)")
    st.write(f"**File:** {res.base_file_name} — **Sheet:** {res.base_sheet_name}")

    # Tampilkan info kedua kolom SO
    col_info1, col_info2 = st.columns(2)
    with col_info1:
        st.success(f"Kolom SO #1 (SAP): **{res.base_col}**\n\n_{res.base_reason}_")
    with col_info2:
        if res.base_col_fvb:
            st.success(f"Kolom SO #2 (FVB): **{res.base_col_fvb}**\n\n_{res.base_reason_fvb}_")
        else:
            st.warning(f"Kolom SO #2 (FVB): Tidak ditemukan\n\n_{res.base_reason_fvb}_")

    st.subheader("✅ Matches")
    st.dataframe(res.matches.head(1000), use_container_width=True)
    st.subheader("❌ Not Found")
    st.dataframe(res.not_found.head(1000), use_container_width=True)
    st.subheader("⚠️ Empty SO")
    st.dataframe(res.empty_so.head(1000), use_container_width=True)

    if not res.log_df.empty:
        st.subheader("🧾 Deteksi Kolom (Log)")
        st.dataframe(res.log_df, use_container_width=True)

    # ===== Drill-down =====
    st.markdown("---")
    st.subheader("🔎 Detail SO yang Match (Drill-down)")

    if res.matched_so_list:
        show_all = st.checkbox("Tampilkan semua SO match (default)", value=True)
        if show_all:
            selected_sos = res.matched_so_list
        else:
            selected_sos = st.multiselect(
                "Pilih SO untuk ditampilkan detailnya:",
                options=res.matched_so_list,
                default=res.matched_so_list[:1]
            )

        if selected_sos:
            with st.expander("Filter tabel Matches ke SO terpilih (opsional)", expanded=False):
                if st.checkbox("Aktifkan filter Matches"):
                    mask = (
                        res.matches["__SO_norm_SAP__"].isin(selected_sos) |
                        res.matches["__SO_norm_FVB__"].isin(selected_sos)
                    )
                    st.dataframe(res.matches[mask], use_container_width=True)

            for so in selected_sos:
                st.markdown(f"### SO: **{so}**")

                # Cari baris di base — cocok di SAP ATAU FVB
                mask_base = (
                    (res.base_df_out["__SO_norm_SAP__"] == so) |
                    (res.base_df_out["__SO_norm_FVB__"] == so)
                )
                base_rows = res.base_df_out[mask_base]

                # Tunjukkan kolom mana yang cocok
                matched_via = []
                if (res.base_df_out["__SO_norm_SAP__"] == so).any():
                    matched_via.append(f"SAP_ODR_NO (`{res.base_col}`)")
                if (res.base_df_out["__SO_norm_FVB__"] == so).any():
                    matched_via.append(f"FVB SO (`{res.base_col_fvb}`)")
                if matched_via:
                    st.caption(f"✅ SO ini ditemukan via: {' & '.join(matched_via)}")

                st.write("**Daily (base) — Rows**")
                if base_rows.empty:
                    st.info("Tidak ada baris di base untuk SO ini.")
                else:
                    st.dataframe(base_rows, use_container_width=True)

                # Baris di referensi
                any_ref = False
                for item in res.ref_tables:
                    df_ref = item["df"]
                    sub = df_ref[df_ref["__SO_norm__"] == so]
                    if not sub.empty:
                        any_ref = True
                        st.write(f"**Reference — {item['file']} · {item['sheet']} (kolom SO: {item['so_col']})**")
                        st.dataframe(sub, use_container_width=True)
                if not any_ref:
                    st.warning("SO ini tidak ditemukan pada tabel referensi yang tersimpan.")
        else:
            st.info("Tidak ada SO terpilih untuk ditampilkan.")
    else:
        st.info("Belum ada SO yang match, jadi detail tidak tersedia.")

    # ==== Download Excel ====
    st.subheader("📥 Download Report (Excel)")
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"SO_AutoDetect_Report_{ts}.xlsx"
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        res.matches.to_excel(writer, index=False, sheet_name="Matches")
        res.not_found.to_excel(writer, index=False, sheet_name="Not_Found")
        res.empty_so.to_excel(writer, index=False, sheet_name="Empty_SO")
        meta = pd.DataFrame({
            "Key": [
                "Base file", "Base sheet",
                "SO column #1 (SAP)", "Reason #1",
                "SO column #2 (FVB)", "Reason #2",
                "Ref files count", "Generated at"
            ],
            "Value": [
                res.base_file_name, res.base_sheet_name,
                res.base_col or "(manual)", res.base_reason,
                res.base_col_fvb or "(tidak ditemukan)", res.base_reason_fvb,
                res.ref_count, ts,
            ],
        })
        meta.to_excel(writer, index=False, sheet_name="Meta")
    output_buffer.seek(0)
    st.download_button(
        label="📥 Download Excel",
        data=output_buffer.getvalue(),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Unggah file & klik **Proses**. Setelah itu, pemilihan SO tidak akan memproses ulang—hanya menampilkan data dari hasil yang sudah disimpan.")
