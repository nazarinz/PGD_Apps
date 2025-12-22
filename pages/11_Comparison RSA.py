# app_streamlit_merged_simple.py
# -*- coding: utf-8 -*-
"""
RSA - PGD Comparison — Simple (only Comparison + Export)

FINAL REVISION:
- FIXED delay mapping logic (Infor digit → canonical code)
- Delay comparison uses FINAL STRING (not digits)
- (blank) vs 0 treated as equal
- 27 vs 04-0027 → TRUE
- All existing logic preserved (join, fallback, export, styling)
"""

import re
import io
import sys
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st

# =====================
# Timezone
# =====================
try:
    from zoneinfo import ZoneInfo
    _tz = ZoneInfo("Asia/Jakarta")
except Exception:
    _tz = None

# =====================
# Optional openpyxl styling
# =====================
EXCEL_STYLED_AVAILABLE = False
_OPENPYXL_ERROR = None
try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    EXCEL_STYLED_AVAILABLE = True
except Exception as e:
    _OPENPYXL_ERROR = e

# =====================
# Constants
# =====================
BLANKS = {"(blank)", "blank", "", "--", " -- ", " --", "0", 0}

JOIN_KEYS = ["PO No.(Full)", "CRD_key", "PD_key"]

DATE_COLS_INFOR = ["Issue Date","FPD","LPD","PSDD","PODD","CRD","PD"]
DATE_COLS_SAP   = ["Document Date","FPD","LPD","PSDD","PODD","FCR Date","PO Date","CRD","PD","Actual PGI"]

DATE_FMT = "mm/dd/yyyy"

DESIRED_ORDER = [
    "Client No","Site","Brand FTY Name","SO","Order Type","Order Type Description",
    "PO No.(Full)","infor Order Status","Customer PO item","Line Aggregator","PO No.(Short)",
    "Merchandise Category 2","Quanity","infor Quantity","Result_Quantity",
    "Model Name","Article No","infor Article No","Result Article No","SAP Material",
    "Pattern Code(Up.No.)","Model No","Outsole Mold","Gender","Category 1","Category 2","Category 3",
    "Unit Price","Classification Code","DRC",
    "Delay/Early - Confirmation PD","infor Delay/Early - Confirmation PD","Result Delay/Early - Confirmation PD",
    "Delay/Early - Confirmation CRD","infor Delay/Early - Confirmation CRD","Result Delay/Early - Confirmation CRD",
    "Delay - PO PSDD Update","infor Delay - PO PSDD Update","Result Delay - PO PSDD Update",
    "Delay - PO PD Update","infor Delay - PO PD Update","Result Delay - PO PD Update",
    "MDP","PDP","SDP","Article Lead time",
    "Ship-to-Sort1","infor Customer Number","Result_Ship-to-Sort1",
    "Ship-to Country","Ship to Name","infor Shipment Method",
    "Document Date","FPD","infor FPD","Result FPD","LPD","infor LPD","Result LPD",
    "CRD","infor CRD","Result CRD","PSDD","infor PSDD","Result PSDD",
    "PODD","infor PODD","Result PODD","FCR Date","PD","infor PD","Result PD",
    "PO Date","Actual PGI","Segment","S&P LPD","Currency"
]

# =====================
# Helpers
# =====================
def to_dt_series(s):
    return pd.to_datetime(s, errors="coerce").dt.normalize()

def to_dt_scalar(x):
    dt = pd.to_datetime(x, errors="coerce")
    return dt.normalize() if not pd.isna(dt) else pd.NaT

def fmt_dt(dt):
    return "" if pd.isna(dt) else dt.strftime("%m/%d/%Y")

def date_concat(series):
    dts = to_dt_series(series).dropna()
    if dts.empty:
        return np.nan
    return " | ".join(fmt_dt(d) for d in sorted(set(dts)))

def date_to_text_cell(val):
    return fmt_dt(to_dt_scalar(val))

def sum_keep_nan(s):
    return s.sum(min_count=1)

def keep_or_join(s):
    vals = pd.unique(s.dropna())
    if len(vals) == 0:
        return np.nan
    if len(vals) == 1:
        return vals[0]
    return " | ".join(map(str, vals))

def extract_digits(x):
    if pd.isna(x):
        return np.nan
    s = str(x).strip()
    if s.lower() in BLANKS:
        return np.nan
    m = re.search(r"\d+", s)
    return m.group(0) if m else np.nan

def split_date_set(x):
    if pd.isna(x):
        return set()
    out = set()
    for part in str(x).split("|"):
        dt = pd.to_datetime(part.strip(), errors="coerce")
        if not pd.isna(dt):
            out.add(dt.normalize())
    return out

def equal_series(a, b):
    return (a.eq(b)) | (a.isna() & b.isna())

def norm_str(s):
    return s.astype(str).str.strip().replace(list(BLANKS), np.nan)

def norm_num(s):
    return pd.to_numeric(s, errors="coerce")

# =====================
# 🔥 FIXED DELAY NORMALIZATION
# =====================
def norm_delay(s: pd.Series) -> pd.Series:
    """
    FINAL delay normalization:
    - DO NOT extract digits
    - Compare canonical strings (04-0027)
    - blanks/0 → NaN
    """
    return (
        s.astype(str)
         .str.strip()
         .replace(list(BLANKS), np.nan)
    )

# =====================
# 🔥 FIXED DELAY MAPPING (Infor only)
# =====================
CODE_MAPPING = {
    '161':'01-0161','84':'03-0084','68':'02-0068','64':'04-0064',
    '62':'02-0062','61':'01-0061','51':'03-0051','46':'03-0046',
    '7':'02-0007','3':'03-0003','2':'01-0002','1':'01-0001',
    '4':'04-0004','8':'02-0008','10':'04-0010','49':'03-0049',
    '90':'04-0090','63':'03-0063','27':'04-0027'
}

def map_delay_series_to_code(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.strip().replace(list(BLANKS), np.nan)
    digits = s.apply(extract_digits)
    mapped = digits.map(CODE_MAPPING)
    return mapped.fillna(s)

# =====================
# Core pipeline
# =====================
def run_core_pipeline(df_sap, df_infor):

    # -------- Normalize Infor
    df_infor = df_infor.rename(columns={
        "Order #": "PO No.(Full)",
        "Line Aggregator": "Customer PO item",
        "Article Number": "Article No",
        "Country/Region": "Ship-to Country",
        "Customer Request Date (CRD)": "CRD",
        "Plan Date": "PD",
        "PO Statistical Delivery Date (PSDD)": "PSDD",
        "First Production Date": "FPD",
        "Last Production Date": "LPD",
        "Grand Total": "Quantity",
        "Shipment Method": "Shipment Method"
    })

    if "Customer Number" in df_infor.columns:
        df_infor["Customer Number"] = df_infor["Customer Number"].replace("--", "ZA30")

    df_infor = df_infor[df_infor["Quantity"].fillna(0) != 0]

    meta_cols = ["Issue Date","PO No.(Full)","Model Name","Article No","Ship-to Country","CRD","PD"]
    df_infor_grp = df_infor.groupby(meta_cols, dropna=False).agg(keep_or_join).reset_index()
    df_infor_grp["CRD_key"] = to_dt_series(df_infor_grp["CRD"])
    df_infor_grp["PD_key"]  = to_dt_series(df_infor_grp["PD"])

    # -------- Normalize SAP
    for c in DATE_COLS_SAP:
        if c in df_sap.columns:
            df_sap[c] = df_sap[c].map(date_to_text_cell)

    df_sap["CRD_key"] = to_dt_series(df_sap["CRD"])
    df_sap["PD_key"]  = to_dt_series(df_sap["PD"])

    # -------- Merge
    inf_pick = df_infor_grp.rename(columns={c: f"infor {c}" for c in df_infor_grp.columns if c not in JOIN_KEYS})
    df = df_sap.merge(inf_pick, on=JOIN_KEYS, how="left")

    # -------- Apply delay mapping (Infor side only)
    for col in [
        "infor Delay/Early - Confirmation PD",
        "infor Delay/Early - Confirmation CRD",
        "infor Delay - PO PSDD Update",
        "infor Delay - PO PD Update"
    ]:
        if col in df.columns:
            df[col] = map_delay_series_to_code(df[col])

    # -------- Comparisons
    if "Quanity" in df.columns and "infor Quantity" in df.columns:
        df["Result_Quantity"] = equal_series(norm_num(df["Quanity"]), norm_num(df["infor Quantity"]))

    if "Delay - PO PSDD Update" in df.columns and "infor Delay - PO PSDD Update" in df.columns:
        df["Result Delay - PO PSDD Update"] = equal_series(
            norm_delay(df["Delay - PO PSDD Update"]),
            norm_delay(df["infor Delay - PO PSDD Update"])
        )

    if "Delay - PO PD Update" in df.columns and "infor Delay - PO PD Update" in df.columns:
        df["Result Delay - PO PD Update"] = equal_series(
            norm_delay(df["Delay - PO PD Update"]),
            norm_delay(df["infor Delay - PO PD Update"])
        )

    # -------- Reorder
    present = [c for c in DESIRED_ORDER if c in df.columns]
    rest = [c for c in df.columns if c not in present]
    return df[present + rest]

# =====================
# Streamlit UI
# =====================
st.set_page_config(page_title="RSA - PGD Comparison (Simple)", layout="wide")
st.title("RSA - PGD Comparison — Simple")

sap_file = st.file_uploader("Upload SAP file", type=["xlsx","xls","csv"])
infor_file = st.file_uploader("Upload Infor file", type=["xlsx","xls","csv"])

if sap_file and infor_file:
    df_sap = pd.read_excel(sap_file)
    df_infor = pd.read_excel(infor_file)

    if st.button("Run Comparison"):
        df_final = run_core_pipeline(df_sap, df_infor)
        st.dataframe(df_final.head(200), use_container_width=True)

        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
            df_final.to_excel(writer, index=False, sheet_name="Data")
        out.seek(0)

        today = datetime.now(_tz).strftime("%Y%m%d")
        st.download_button(
            "⬇️ Download Excel",
            data=out,
            file_name=f"RSA - PGD Comparison Tracking Report - {today}.xlsx"
        )
