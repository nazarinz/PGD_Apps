# PGD Apps — Input Tracking Report (Pending Cancel Only)
# =====================================================
# Sub-tools: "Buat masukin tracking PO Pending Cancel"
# - Sheet 1: hasil filter + reorder lengkap (TARGET_ORDER)
# - Sheet 2: hasil ringkas (SLIM_ORDER)
# - Normalisasi minimal: "Sales Order" -> "SO", "Article Lead Time" -> "Lead Time"
# - Kolom hilang dibuat NA agar urutan tetap konsisten

import re
import pandas as pd
import streamlit as st
from utils import set_page, header, footer, write_excel_autofit

# =========================
# Streamlit Page
# =========================
set_page("PGD Apps — Input Tracking Report (Pending Cancel)", "📝")
header("📝 Input Tracking — Pending Cancel")

st.caption("Sub-tools ini memfilter & menyusun ulang kolom untuk tracking PO Pending Cancel. "
           "Output terdiri dari 2 sheet: Sheet 1 (lengkap) dan Sheet 2 (ringkas).")

# =========================
# Target urutan kolom (Sheet 1)
# =========================
TARGET_ORDER = [
    "Working Status","Select","Working Status Descr.","Requirement Segment","SO",
    "Sold-To PO No.","CRD","CRD-DRC","PD","POSDD-DRC","POSDD","FPD","FPD-DRC",
    "PODD","PODD-DRC","Est. Inspection Date","LPD","LPD-DRC","FGR","Cust Article No.",
    "Model Name","Article","Lead Time","Season","Product Hierarchy 3","Ship-To Search Term",
    "Ship-To Country","Document Date","Order Quantity","Order Type"
]

# =========================
# Target urutan kolom (Sheet 2 - ringkas)
# =========================
SLIM_ORDER = [
    "Sales Order","Sold-To PO No.","Cust Article No.","Model Name",
    "Ship-To Search Term","Ship-To Country","Document Date","Order Quantity","Order Type"
]

# =========================
# Normalisasi header (minimal)
# =========================
def _normalize_header(name: str) -> str:
    base = re.sub(r"\s+", " ", str(name).strip())
    alias_map = {
        # Identity / common
        "Working Status": "Working Status",
        "Select": "Select",
        "Working Status Descr.": "Working Status Descr.",
        "Requirement Segment": "Requirement Segment",
        "SO": "SO",
        "Sold-To PO No.": "Sold-To PO No.",
        "CRD": "CRD",
        "CRD-DRC": "CRD-DRC",
        "PD": "PD",
        "POSDD-DRC": "POSDD-DRC",
        "POSDD": "POSDD",
        "FPD": "FPD",
        "FPD-DRC": "FPD-DRC",
        "PODD": "PODD",
        "PODD-DRC": "PODD-DRC",
        "Est. Inspection Date": "Est. Inspection Date",
        "LPD": "LPD",
        "LPD-DRC": "LPD-DRC",
        "FGR": "FGR",
        "Cust Article No.": "Cust Article No.",
        "Model Name": "Model Name",
        "Article": "Article",
        "Lead Time": "Lead Time",
        "Season": "Season",
        "Product Hierarchy 3": "Product Hierarchy 3",
        "Ship-To Search Term": "Ship-To Search Term",
        "Ship-To Country": "Ship-To Country",
        "Document Date": "Document Date",
        "Order Quantity": "Order Quantity",
        "Order Type": "Order Type",
        "Sales Order": "Sales Order",
        # Aliases → target (yang diperlukan saja)
        "Sales Order": "Sales Order",   # tetap
        "Article Lead Time": "Lead Time",
    }
    return alias_map.get(base, base)

def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [_normalize_header(c) for c in df.columns]
    return df

def filter_and_reorder(df: pd.DataFrame, target_cols: list) -> tuple[pd.DataFrame, list]:
    # Normalisasi header dulu agar alias dikenali (Sales Order→Sales Order, SO tetap, Article Lead Time→Lead Time)
    df = _normalize_columns(df).copy()
    missing = [c for c in target_cols if c not in df.columns]
    for col in missing:
        df[col] = pd.NA
    out_df = df[target_cols]
    return out_df, missing

# =========================
# UI — Urutan kolom & Upload
# =========================
with st.expander("📋 Urutan kolom output (Sheet 1)", expanded=False):
    st.dataframe(pd.DataFrame({"Urutan Kolom Sheet 1": TARGET_ORDER}), use_container_width=True)

with st.expander("📋 Urutan kolom output (Sheet 2 - Ringkas)", expanded=False):
    st.dataframe(pd.DataFrame({"Urutan Kolom Sheet 2": SLIM_ORDER}), use_container_width=True)

uploaded = st.file_uploader("Upload Excel (.xlsx/.xls)", type=["xlsx","xls"], accept_multiple_files=False)
if not uploaded:
    st.stop()

# =========================
# Proses semua sheet
# =========================
xls = pd.read_excel(uploaded, sheet_name=None, engine="openpyxl")
out_frames = {}     # Sheet 1 (per sheet)
slim_frames = []    # akumulasi untuk Sheet 2 (gabungan semua sheet)
report = []

for sheet_name, df in xls.items():
    # Sheet 1
    df_out, missing = filter_and_reorder(df, TARGET_ORDER)
    out_frames[sheet_name] = df_out

    # Sheet 2 (ringkas)
    slim_tmp = _normalize_columns(df).copy()

    # Pastikan 'Sales Order' (bukan SO) tersedia untuk sheet 2
    if 'Sales Order' not in slim_tmp.columns:
        # kalau kolom 'SO' ada, gunakan sebagai 'Sales Order'
        if 'SO' in slim_tmp.columns:
            slim_tmp['Sales Order'] = slim_tmp['SO']
        else:
            slim_tmp['Sales Order'] = pd.NA

    for col in SLIM_ORDER:
        if col not in slim_tmp.columns:
            slim_tmp[col] = pd.NA

    slim_view = slim_tmp[SLIM_ORDER].copy()
    slim_view.insert(0, 'Source Sheet', sheet_name)  # opsional: jejak asal
    slim_frames.append(slim_view)

    report.append({
        "Sheet": sheet_name,
        "Kolom hilang Sheet1 (dibuat NA)": ", ".join(missing) if missing else "-",
        "Rows Sheet1": len(df_out),
        "Cols Sheet1": len(df_out.columns),
        "Rows Sheet2": len(slim_view),
    })

report_df = pd.DataFrame(report).sort_values("Sheet")
st.success("✅ Berhasil diproses. Cek ringkasan & unduh hasilnya.")
st.dataframe(report_df, use_container_width=True)

# =========================
# Export Excel
# =========================
if len(slim_frames) > 0:
    sheet2_df = pd.concat(slim_frames, ignore_index=True)
    payload = write_excel_autofit({**out_frames, "Sheet2_PendingCancel_Slim": sheet2_df, "Report": report_df})
else:
    payload = write_excel_autofit({**out_frames, "Report": report_df})

st.download_button(
    "⬇️ Download Hasil Pending Cancel (Sheet1 & Sheet2)",
    data=payload,
    file_name="pending_cancel_filtered.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

footer()
