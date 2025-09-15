# Buat masukin tracking PO Pending Cancel

# ============================
# Reorder Kolom Excel (All Sheets)
# - Colab: upload file
# - Lokal: input path
# ============================

import pandas as pd
import os
import streamlit as st
from utils import set_page, header, footer, write_excel_autofit

# --- Urutan kolom yang diinginkan ---
DESIRED_ORDER = [
    "Working Status","Select","Working Status Descr.","Requirement Segment",
    "Sales Order","Sold-To PO No.","CRD","CRD-DRC","PD","POSDD-DRC","POSDD",
    "FPD","FPD-DRC","PODD","PODD-DRC","Est. Inspection Date",
    "LPD","LPD-DRC","FGR","Cust Article No.","Model Name","Article",
    "Lead Time","Season","Product Hierarchy 3","Ship-To Search Term",
    "Ship-To Country","Document Date","Order Quantity","Order Type"
]

# --- Kolom target & urutan yang diinginkan ---
TARGET_ORDER = [
    "Sales Order","Remark","Sold-To PO No.","Ship-To Party PO No.",
    "Model Name","Cust Article No.","Article","Order Quantity",
    "CRD","PD","LPD","PODD","Ship-To Search Term","Ship-To Name"
]

def reorder_columns(df: pd.DataFrame, desired_order: list) -> tuple[pd.DataFrame, list, list]:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in desired_order if c not in df.columns]
    extra   = [c for c in df.columns if c not in desired_order]
    for col in missing:
        df[col] = pd.NA
    new_order = desired_order + [c for c in df.columns if c not in desired_order]
    df = df[new_order]
    return df, missing, extra

def filter_and_reorder(df: pd.DataFrame, target_cols: list) -> tuple[pd.DataFrame, list]:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in target_cols if c not in df.columns]
    for col in missing:
        df[col] = pd.NA
    out_df = df[target_cols]
    return out_df, missing

# --- Streamlit Page ---
set_page("PGD Apps — Input Tracking Report", "📝")
header("📝 Input Tracking Report")
choice = st.selectbox("Pilih sub-tools", [
    "Buat masukin trackingan PO Reroute",
    "Buat masukin tracking PO Pending Cancel",
])

if choice == "Buat masukin trackingan PO Reroute":
    up = st.file_uploader("Upload Excel (.xlsx/.xls)", type=["xlsx","xls"], accept_multiple_files=False, key="itr_reorder")
    if not up:
        st.stop()
    xls = pd.read_excel(up, sheet_name=None, engine="openpyxl")
    out_frames = {}
    report = []
    for sheet_name, df in xls.items():
        df_out, missing, extra = reorder_columns(df, DESIRED_ORDER)
        out_frames[sheet_name] = df_out
        report.append({
            "Sheet": sheet_name,
            "Missing": ", ".join(missing) if missing else "-",
            "Extra": ", ".join(extra) if extra else "-",
        })
    report_df = pd.DataFrame(report)
    payload = write_excel_autofit({**out_frames, "Report": report_df})
    st.dataframe(report_df, use_container_width=True)
    st.download_button("⬇️ Download Hasil Reorder", data=payload, file_name="reordered_output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if choice == "Buat masukin tracking PO Pending Cancel":
    up2 = st.file_uploader("Upload Excel (.xlsx/.xls)", type=["xlsx","xls"], accept_multiple_files=False, key="itr_filter")
    if not up2:
        st.stop()
    xls = pd.read_excel(up2, sheet_name=None, engine="openpyxl")
    out_frames = {}
    report = []
    for sheet_name, df in xls.items():
        df_out, missing = filter_and_reorder(df, TARGET_ORDER)
        out_frames[sheet_name] = df_out
        report.append({
            "Sheet": sheet_name,
            "MissingAdded": ", ".join(missing) if missing else "-",
        })
    report_df = pd.DataFrame(report)
    payload = write_excel_autofit({**out_frames, "Report": report_df})
    st.dataframe(report_df, use_container_width=True)
    st.download_button("⬇️ Download Hasil Filter+Reorder", data=payload, file_name="filtered_output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

footer()
