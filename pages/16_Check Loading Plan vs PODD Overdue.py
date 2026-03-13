from utils.auth import require_login

require_login()

import streamlit as st
import pandas as pd
import io
import re
from collections import defaultdict
from datetime import datetime
import traceback

st.set_page_config(page_title="Loading Plan Checker", page_icon="📋", layout="wide")


# =========================================================
# STANDARD LOADING PLAN READER (Header baris ke-3)
# =========================================================
def read_daily_loading_plan(uploaded_file):
    name = uploaded_file.name.lower()
    ext = "." + name.split(".")[-1]

    if hasattr(uploaded_file, "getvalue"):
        raw = uploaded_file.getvalue()
    else:
        uploaded_file.seek(0)
        raw = uploaded_file.read()

    if ext in [".xlsx", ".xlsm"]:
        engine = "openpyxl"
    elif ext == ".xls":
        engine = "xlrd"
    elif ext == ".xlsb":
        engine = "pyxlsb"
    elif ext == ".ods":
        engine = "odf"
    else:
        raise ValueError(f"Format {ext} tidak didukung")

    sheets = pd.read_excel(io.BytesIO(raw), sheet_name=None, header=None, engine=engine)

    sheet_name = None
    for sh in sheets.keys():
        if str(sh).strip().lower().replace(" ", "_") == "loading_plan":
            sheet_name = sh
            break
    if sheet_name is None:
        sheet_name = list(sheets.keys())[0]

    df_raw = sheets[sheet_name]

    header_row_index = 2
    raw_header = [str(x).strip() if x is not None else "" for x in df_raw.iloc[header_row_index].tolist()]

    seen = set()
    clean_cols = []
    for c in raw_header:
        base = c if c else "col"
        new = base
        i = 1
        while new in seen:
            i += 1
            new = f"{base}_{i}"
        seen.add(new)
        clean_cols.append(new)

    df = df_raw.iloc[header_row_index + 1:].copy()
    df.columns = clean_cols
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")
    df = df.reset_index(drop=True)

    return df, sheet_name


# =========================================================
# STANDARD SO DETECTOR
# =========================================================
CANDIDATES = [
    "sap_odr_no", "sap_odrno", "sap_order_no", "sap_odr_number",
    "sales_order", "salesorder", "so", "so_no", "sono", "so_number",
]

FVB_CANDIDATES = [
    "fvb_so", "fvbso", "fvb_so", "fvb_sales_order", "fvb",
]

def norm_col_name(c: str) -> str:
    c = str(c).strip().lower()
    c = re.sub(r"\s+", " ", c)
    c = re.sub(r"[^0-9a-z_ ]+", "", c)
    c = c.replace(" ", "_")
    c = re.sub(r"_+", "_", c)
    return c

def detect_so_column(df: pd.DataFrame):
    cols = [norm_col_name(c) for c in df.columns]

    for cand in CANDIDATES:
        if cand in cols:
            return df.columns[cols.index(cand)]

    for i, cn in enumerate(cols):
        if ("sap" in cn) and ("odr" in cn or "order" in cn):
            return df.columns[i]
        if ("sales" in cn) and ("order" in cn):
            return df.columns[i]

    return None


def detect_fvb_so_column(df: pd.DataFrame):
    """Deteksi kolom FVB SO. Return nama kolom asli atau None."""
    cols_norm = [norm_col_name(c) for c in df.columns]

    # 1) Eksak dari kandidat FVB
    for cand in FVB_CANDIDATES:
        nc = norm_col_name(cand)
        if nc in cols_norm:
            return df.columns[cols_norm.index(nc)]

    # 2) Heuristik: nama kolom mengandung "fvb"
    for i, cn in enumerate(cols_norm):
        if "fvb" in cn:
            return df.columns[i]

    return None


def clean_so_series(series: pd.Series) -> pd.Series:
    """Normalisasi nilai SO: strip, hapus .0, filter angka saja."""
    return (
        series
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )


# =========================================================
def extract_date_from_filename_ddmm(filename):
    name = filename.rsplit(".", 1)[0]
    try:
        parts = name.split(".")
        if len(parts) != 2:
            return None
        day, month = map(int, parts)
        today = pd.Timestamp.today()
        year = today.year
        if month > today.month:
            year -= 1
        return pd.Timestamp(year=year, month=month, day=day)
    except:
        return None


# =========================================================
def process_files(zrsd_file, plan_files, date_format):
    text_columns = {
        'PO No.(Full)': str,
        'PO No.(Short)': str,
        'Article No': str,
        'SAP Material': str,
        'Pattern Code(Up.No.)': str,
        'Model No': str,
        'Outsole Mold': str,
        'SO': str,
        'Material': str
    }

    df_013 = pd.read_excel(io.BytesIO(zrsd_file.getvalue()), dtype=text_columns)

    fcr_date_check = pd.to_datetime(df_013['FCR Date'], errors='coerce').isna()
    podd_check = pd.to_datetime(df_013['PODD'], errors='coerce')

    today = pd.Timestamp.today().normalize()
    end_date = today + pd.Timedelta(days=3)

    df_filtered = df_013[fcr_date_check & (podd_check <= end_date)].copy()

    known_date_columns = [
        'Document Date', 'FPD', 'LPD', 'CRD', 'PSDD', 'FCR Date',
        'PODD', 'PD', 'PO Date', 'Actual PGI', 'Delivery Date',
        'Ship Date', 'Created Date', 'Modified Date', 'Invoice Date'
    ]

    for col in df_filtered.columns:
        if any(col.lower() == known.lower() for known in known_date_columns):
            if not pd.api.types.is_datetime64_any_dtype(df_filtered[col]):
                df_filtered[col] = pd.to_datetime(df_filtered[col], errors='coerce')

    # plan_so_map: SO (int) → list of dates
    # SO bisa berasal dari kolom SAP_ODR_NO maupun FVB SO
    plan_so_map = defaultdict(list)

    # ===================== READ LOADING PLAN =====================
    for file in plan_files:
        try:
            st.info(f"📂 Membaca {file.name}...")

            df_plan, sheet = read_daily_loading_plan(file)
            plan_date = extract_date_from_filename_ddmm(file.name)

            if plan_date is None:
                st.error(f"❌ {file.name}: Format nama file salah (harus DD.MM seperti 22.01)")
                continue

            # ── Kolom SO #1: SAP_ODR_NO ──────────────────────────────────
            so_col_sap = detect_so_column(df_plan)
            if not so_col_sap:
                st.error(f"❌ {file.name}: Kolom SAP SO tidak ditemukan")
                st.write("Kolom yang tersedia:", list(df_plan.columns))
                continue

            # ── Kolom SO #2: FVB SO ───────────────────────────────────────
            so_col_fvb = detect_fvb_so_column(df_plan)

            # Log info kolom yang ditemukan
            fvb_label = f"`{so_col_fvb}`" if so_col_fvb else "tidak ditemukan"
            st.success(
                f"✅ {file.name} → SAP SO: `{so_col_sap}` | "
                f"FVB SO: {fvb_label} | Tanggal: {plan_date.date()}"
            )

            # ── Masukkan SO dari kolom SAP ke map ─────────────────────────
            df_plan["__SO_sap__"] = clean_so_series(df_plan[so_col_sap])
            for so in df_plan["__SO_sap__"].dropna().unique():
                if str(so).isdigit():
                    plan_so_map[int(so)].append(("SAP", plan_date.date()))

            # ── Masukkan SO dari kolom FVB ke map (jika ada) ──────────────
            if so_col_fvb and so_col_fvb in df_plan.columns:
                df_plan["__SO_fvb__"] = clean_so_series(df_plan[so_col_fvb])
                for so in df_plan["__SO_fvb__"].dropna().unique():
                    if str(so).isdigit():
                        plan_so_map[int(so)].append(("FVB", plan_date.date()))

        except Exception as e:
            st.error(f"❌ {file.name}: {str(e)}")
            with st.expander("🔍 Detail Error"):
                st.code(traceback.format_exc())
            continue

    # ===================== CHECK RESULT =====================
    def check_remark(row):
        try:
            so_raw = str(row['SO']).strip()
            so = int(so_raw.replace(".0", ""))
            podd_date = pd.to_datetime(row['PODD'], errors='coerce').date()
        except:
            return "⚠️ Invalid Data", "mismatch", "", ""

        if so not in plan_so_map:
            return "❌ NOT IN LOADING PLAN", "not_found", "", ""

        entries = plan_so_map[so]  # list of (source, date)

        # Pisahkan tanggal per sumber
        sap_dates  = sorted({d for src, d in entries if src == "SAP"})
        fvb_dates  = sorted({d for src, d in entries if src == "FVB"})
        all_dates  = sorted({d for _, d in entries})

        sap_dates_str = ", ".join(str(d) for d in sap_dates) if sap_dates else "-"
        fvb_dates_str = ", ".join(str(d) for d in fvb_dates) if fvb_dates else "-"
        plan_dates_str = f"SAP: {sap_dates_str} | FVB: {fvb_dates_str}"

        # Tentukan sumber mana yang menyimpan SO ini
        sources_found = []
        if sap_dates:
            sources_found.append("SAP")
        if fvb_dates:
            sources_found.append("FVB")
        source_label = "+".join(sources_found)

        if podd_date in all_dates:
            return (
                f"✅ MATCH – Date Match (via {source_label})",
                "match",
                plan_dates_str,
                source_label,
            )
        else:
            return (
                f"⚠️ IN PLAN – Date Mismatch (Plan: {plan_dates_str})",
                "mismatch",
                plan_dates_str,
                source_label,
            )

    result_df = df_filtered.apply(
        lambda row: pd.Series(check_remark(row)), axis=1
    )

    df_filtered['Remark Loading Plan'] = result_df[0]
    df_filtered['Status']              = result_df[1]
    df_filtered['Plan Dates']          = result_df[2]
    df_filtered['SO Source']           = result_df[3]   # kolom baru: SAP / FVB / SAP+FVB

    return df_filtered


# =========================================================
# ======================= UI ================================
# =========================================================
st.title("📋 Loading Plan Checker")

st.sidebar.header("⚙️ Pengaturan")
st.sidebar.markdown("---")

with st.sidebar:
    zrsd_file = st.file_uploader("Upload ZRSD", type=["xlsx", "xls"])
    plan_files = st.file_uploader(
        "Upload Loading Plan (DD.MM format)",
        type=["ods", "xlsx", "xls"],
        accept_multiple_files=True
    )

    if st.button("🚀 Proses Data", use_container_width=True):
        if zrsd_file and plan_files:
            with st.spinner("Memproses data..."):
                st.session_state['results'] = process_files(zrsd_file, plan_files, "dummy")
        else:
            st.warning("⚠️ Upload semua file yang diperlukan")

# =========================================================
if 'results' in st.session_state:
    df = st.session_state['results']

    st.subheader("📊 Hasil Pengecekan")
    st.dataframe(df, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

        from openpyxl.styles import numbers
        worksheet = writer.sheets['Sheet1']

        known_date_columns = [
            'Document Date', 'FPD', 'LPD', 'CRD', 'PSDD', 'FCR Date',
            'PODD', 'PD', 'PO Date', 'Actual PGI', 'Delivery Date',
            'Ship Date', 'Created Date', 'Modified Date', 'Invoice Date'
        ]

        date_col_indices = []
        for i, col in enumerate(df.columns, start=1):
            if any(col.lower() == known.lower() for known in known_date_columns):
                date_col_indices.append(i)

        for col_idx in date_col_indices:
            col_letter = worksheet.cell(row=1, column=col_idx).column_letter
            for row in range(2, worksheet.max_row + 1):
                cell = worksheet[f'{col_letter}{row}']
                if cell.value:
                    cell.number_format = 'M/D/YYYY'

    st.download_button(
        "📥 Download Excel",
        output.getvalue(),
        file_name=f"loading_plan_result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        use_container_width=True
    )
