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
        # xlrd old .xls; jika gagal, coba openpyxl
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
    base_col: str
    base_reason: str
    ref_count: int
    matches: pd.DataFrame
    not_found: pd.DataFrame
    empty_so: pd.DataFrame
    base_df_out: pd.DataFrame        # punya kolom __SO_norm__
    ref_tables: list                 # list of dict: {file, sheet, so_col, df}
    log_df: pd.DataFrame
    matched_so_list: list

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

    # Baca ulang tanpa header untuk reheader baris ke-3
    if hasattr(base_file, "getvalue"):
        raw_bytes = base_file.getvalue()
    else:
        try:
            base_file.seek(0)
        except Exception:
            pass
        raw_bytes = base_file.read()

    ext = "." + base_file.name.split(".")[-1].lower()
    if ext == ".ods":
        engine = "odf"
    elif ext in (".xlsx", ".xlsm"):
        engine = "openpyxl"
    elif ext == ".xls":
        engine = "xlrd"
    elif ext == ".xlsb":
        engine = "pyxlsb"
    else:
        engine = "openpyxl"

    book_headerless = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=None, engine=engine, header=None)
    _base_df_raw = book_headerless[base_sheet_name]
    base_df = reheader_with_row(_base_df_raw, header_row_index=2)

    # Deteksi kolom SO
    base_col = pick_column(base_df, "SAP_ODR_NO")
    base_reason = "Forced: SAP_ODR_NO | header=baris 3"
    if not base_col:
        base_col, base_reason = detect_so_column(base_df)
        base_reason = f"{base_reason} | header=baris 3"

    base_df_out = base_df.copy()
    if base_col is None:
        base_col = base_df_out.columns[0]
        base_reason = f"Fallback: pakai kolom pertama ({base_col})"
    base_df_out["__SO_norm__"] = normalize_so_series(base_df_out[base_col])

    # Kumpulkan SO referensi + simpan tabel referensi
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
                if not col:
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
                    "df": df2,  # sudah ada __SO_norm__
                })

    # Hasil utama
    out = base_df_out.copy()
    out["Found_in_reference"] = out["__SO_norm__"].apply(lambda x: (x in all_ref_sos) if pd.notna(x) else False)
    out["Source_Files"] = out["__SO_norm__"].apply(
        lambda x: ", ".join(sorted(so_to_files.get(x, []))) if pd.notna(x) and x in so_to_files else ""
    )
    matches   = out[out["Found_in_reference"] == True].copy()
    not_found = out[(out["__SO_norm__"].notna()) & (out["Found_in_reference"] == False)].copy()
    empty_so  = out[out["__SO_norm__"].isna()].copy()
    matched_so_list = sorted(matches["__SO_norm__"].dropna().astype(str).unique().tolist())

    log_df = pd.DataFrame(log_rows, columns=["File", "Sheet", "Reason"]) if log_rows else pd.DataFrame(columns=["File","Sheet","Reason"])

    return SOResult(
        base_file_name = base_file.name,
        base_sheet_name = base_sheet_name,
        base_col = base_col,
        base_reason = base_reason,
        ref_count = len(ref_files or []),
        matches = matches,
        not_found = not_found,
        empty_so = empty_so,
        base_df_out = base_df_out,
        ref_tables = ref_tables,
        log_df = log_df,
        matched_so_list = matched_so_list
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

# Tombol
colA, colB = st.columns([1,1])
with colA:
    do_process = st.button("▶️ Proses", type="primary")
with colB:
    do_reset = st.button("♻️ Reset hasil")

if do_reset:
    st.session_state.pop("so_detect_result", None)
    st.success("Hasil dipulihkan. Silakan klik Proses lagi setelah pilih file.")

# Jalankan proses hanya saat klik Proses
if do_process:
    if not base_file:
        st.error("Unggah dulu **Daily file**.")
        st.stop()
    st.session_state["so_detect_result"] = process_files(base_file, ref_files)

# =========================
# Render dari cache (tanpa proses ulang)
# =========================
if "so_detect_result" in st.session_state:
    res: SOResult = st.session_state["so_detect_result"]

    st.subheader("📄 Daily file (base)")
    st.write(f"**File:** {res.base_file_name} — **Sheet:** {res.base_sheet_name}")
    st.success(f"Kolom SO terdeteksi: **{res.base_col}** ({res.base_reason})")

    st.subheader("✅ Matches")
    st.dataframe(res.matches.head(1000), use_container_width=True)
    st.subheader("❌ Not Found")
    st.dataframe(res.not_found.head(1000), use_container_width=True)
    st.subheader("⚠️ Empty SO")
    st.dataframe(res.empty_so.head(1000), use_container_width=True)

    if not res.log_df.empty:
        st.subheader("🧾 Deteksi Kolom (Log)")
        st.dataframe(res.log_df, use_container_width=True)

    # ===== Drill-down: DEFAULT tampilkan SEMUA SO match =====
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
                    st.dataframe(res.matches[res.matches["__SO_norm__"].isin(selected_sos)], use_container_width=True)

            for so in selected_sos:
                st.markdown(f"### SO: **{so}**")

                # 1) Baris di base
                base_rows = res.base_df_out[res.base_df_out["__SO_norm__"] == so]
                st.write("**Daily (base) — Rows**")
                if base_rows.empty:
                    st.info("Tidak ada baris di base untuk SO ini.")
                else:
                    st.dataframe(base_rows, use_container_width=True)

                # 2) Baris di setiap referensi
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
            "Key": ["Base file", "Base sheet", "SO column (base)", "Reason", "Ref files count", "Generated at"],
            "Value": [res.base_file_name, res.base_sheet_name, res.base_col or "(manual)", res.base_reason, res.ref_count, ts],
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
