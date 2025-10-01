# PGD Apps — Input Tracking Report (Pending Cancel Only)
# =====================================================
# Sub-tools: "Buat masukin tracking PO Pending Cancel"
# Revisi utama:
# - Menyusun ulang kolom sesuai urutan yang diminta (lihat TARGET_ORDER)
# - Normalisasi header seminimal mungkin (tanpa alias opsional Part Number/Shipping Type)
# - Alias yang diterapkan: "Sales Order" -> "SO" dan "Article Lead Time" -> "Lead Time"
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

st.caption("Sub-tools ini memfilter & menyusun ulang kolom untuk tracking PO Pending Cancel sesuai format yang kamu minta.")

# =========================
# Target urutan kolom (sesuai permintaan)
# =========================
TARGET_ORDER = [
    "Working Status","Select","Working Status Descr.","Requirement Segment","SO",
    "Sold-To PO No.","CRD","CRD-DRC","PD","POSDD-DRC","POSDD","FPD","FPD-DRC",
    "PODD","PODD-DRC","Est. Inspection Date","LPD","LPD-DRC","FGR","Cust Article No.",
    "Model Name","Article","Lead Time","Season","Product Hierarchy 3","Ship-To Search Term",
    "Ship-To Country","Document Date","Order Quantity","Order Type"
]

# =========================
# Normalisasi header (minimal, sesuai kebutuhan)
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
        # Aliases → target (hanya yang diperlukan)
        "Sales Order": "SO",
        "Article Lead Time": "Lead Time",
    }
    return alias_map.get(base, base)


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [_normalize_header(c) for c in df.columns]
    return df


def filter_and_reorder(df: pd.DataFrame, target_cols: list) -> tuple[pd.DataFrame, list]:
    # Normalisasi header dulu agar alias dikenali (Sales Order→SO, Article Lead Time→Lead Time)
    df = _normalize_columns(df).copy()
    missing = [c for c in target_cols if c not in df.columns]
    for col in missing:
        df[col] = pd.NA
    out_df = df[target_cols]
    return out_df, missing

# =========================
# UI — Pemrosesan
# =========================
with st.expander("📋 Urutan kolom output (tetap)", expanded=False):
    st.dataframe(pd.DataFrame({"Urutan Kolom": TARGET_ORDER}), use_container_width=True)

uploaded = st.file_uploader("Upload Excel (.xlsx/.xls)", type=["xlsx","xls"], accept_multiple_files=False)
if not uploaded:
    st.stop()

# Baca semua sheet
xls = pd.read_excel(uploaded, sheet_name=None, engine="openpyxl")
out_frames = {}
report = []

for sheet_name, df in xls.items():
    df_out, missing = filter_and_reorder(df, TARGET_ORDER)
    out_frames[sheet_name] = df_out
    report.append({
        "Sheet": sheet_name,
        "Kolom hilang (dibuat NA)": ", ".join(missing) if missing else "-",
        "Rows": len(df_out),
        "Cols": len(df_out.columns),
    })

report_df = pd.DataFrame(report).sort_values("Sheet")
st.success("✅ Berhasil diproses. Cek ringkasan & unduh hasilnya.")
st.dataframe(report_df, use_container_width=True)

# Export Excel (pakai util project)
payload = write_excel_autofit({**out_frames, "Report": report_df})
st.download_button(
    "⬇️ Download Hasil Pending Cancel",
    data=payload,
    file_name="pending_cancel_filtered.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

footer()
