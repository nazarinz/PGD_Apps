# app_streamlit_merged_simple.py
# -*- coding: utf-8 -*-
"""
RSA - PGD Comparison — Simple (FULL FINAL, STYLED)

- SAP TIDAK digabung (row-level)
- JOIN via PO + CRD_key + PD_key
- Infor digroup
- Delay mapping FIXED
- Ship-to-Sort1 vs Customer Number
- Shortdate Excel (filterable)
- Styled export (openpyxl) + xlsxwriter
"""

import re
import io
import sys
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st

# =========================
# TIMEZONE
# =========================
try:
    from zoneinfo import ZoneInfo
    _tz = ZoneInfo("Asia/Jakarta")
except Exception:
    _tz = None

# =========================
# OPENPYXL (STYLING)
# =========================
EXCEL_STYLED_AVAILABLE = False
_OPENPYXL_ERROR = None
try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    EXCEL_STYLED_AVAILABLE = True
except Exception as e:
    _OPENPYXL_ERROR = e

# =========================
# CONSTANTS
# =========================
JOIN_KEYS = ["PO No.(Full)", "CRD_key", "PD_key"]

DATE_COLS_SAP = [
    "Document Date","FPD","LPD","PSDD","PODD",
    "FCR Date","PO Date","CRD","PD","Actual PGI"
]

DATE_COLS_INFOR = ["Issue Date","FPD","LPD","PSDD","PODD","CRD","PD"]

BLANKS = {"(blank)", "blank", "", "--", " -- ", " --"}

DATE_FMT_XLSX = "mm/dd/yyyy"
DATE_FMT_OPENPYXL = "m/d/yyyy"

# HEADER COLORS (FIXED AS REQUESTED)
INFOR_COLOR  = "FFF9F16D"   # yellow
RESULT_COLOR = "FFC6EFCE"   # green
OTHER_COLOR  = "FFD9D9D9"   # grey

# =========================
# HELPERS
# =========================
def to_dt_series(s):
    return pd.to_datetime(s, errors="coerce").dt.normalize()

def to_dt_scalar(x):
    ts = pd.to_datetime(x, errors="coerce")
    if pd.isna(ts):
        return pd.NaT
    return ts.normalize()

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
    parts = [p.strip() for p in str(x).split("|")]
    out = set()
    for p in parts:
        dt = pd.to_datetime(p, errors="coerce")
        if not pd.isna(dt):
            out.add(dt.normalize())
    return out

def equal_series(a, b):
    return a.eq(b) | (a.isna() & b.isna())

# =========================
# DELAY CODE MAPPING (FIXED)
# =========================
CODE_MAPPING = {
    '161':'01-0161','84':'03-0084','68':'02-0068','64':'04-0064',
    '62':'02-0062','61':'01-0061','51':'03-0051','46':'03-0046',
    '7':'02-0007','3':'03-0003','2':'01-0002','1':'01-0001',
    '4':'04-0004','8':'02-0008','10':'04-0010',
    '49':'03-0049','90':'04-0090','63':'03-0063',
    '27':'04-0027'   # ⬅️ FIX kasus kamu
}

def map_delay_series_to_code(s):
    base = s.apply(extract_digits)
    mapped = base.map(CODE_MAPPING)
    return mapped.where(mapped.notna(), base)

# =========================
# CORE PIPELINE
# =========================
def run_core_pipeline(df_sap_raw, df_infor_raw):

    df_sap = df_sap_raw.copy()
    df_infor = df_infor_raw.copy()

    today = datetime.now(_tz) if _tz else datetime.now()
    out_name = f"RSA - PGD Comparison Tracking Report - {today:%Y%m%d}.xlsx"

    # ---------- INFOR NORMALIZE ----------
    rename_inf = {
        "Order #":"PO No.(Full)",
        "Line Aggregator":"Customer PO item",
        "Article Number":"Article No",
        "Country/Region":"Ship-to Country",
        "Customer Request Date (CRD)":"CRD",
        "Plan Date":"PD",
        "PO Statistical Delivery Date (PSDD)":"PSDD",
        "First Production Date":"FPD",
        "Last Production Date":"LPD",
        "Grand Total":"Quantity",
        "Delivery Delay Pd":"Delay - PO PD Update",
        "Delay - PO Del Update":"Delay - PO PSDD Update",
        "Delay - Confirmation":"Delay/Early - Confirmation CRD",
    }
    df_infor.rename(columns=rename_inf, inplace=True)

    if "Customer Number" in df_infor.columns:
        df_infor["Customer Number"] = df_infor["Customer Number"].replace("--","ZA30")

    df_infor = df_infor[df_infor["Quantity"].fillna(0) != 0].copy()

    meta_cols = ["Issue Date","PO No.(Full)","Model Name","Article No","Ship-to Country","CRD","PD"]

    size_pat = re.compile(r'^(?:[1-9]|1[0-8])(?:K|-K|-)?$')
    size_cols = [c for c in df_infor.columns if size_pat.match(str(c))]

    sum_cols = size_cols + ["Quantity"]
    other_cols = [c for c in df_infor.columns if c not in meta_cols + sum_cols]

    agg = {c:sum_keep_nan for c in sum_cols}
    for c in other_cols:
        agg[c] = date_concat if c in DATE_COLS_INFOR else keep_or_join

    df_infor_g = df_infor.groupby(meta_cols, dropna=False).agg(agg).reset_index()

    df_infor_g["Line Aggregator"] = df_infor_g.get("Customer PO item")

    df_infor_g["CRD_key"] = to_dt_series(df_infor_g["CRD"])
    df_infor_g["PD_key"]  = to_dt_series(df_infor_g["PD"])

    # ---------- SAP NORMALIZE ----------
    for c in DATE_COLS_SAP:
        if c in df_sap.columns:
            df_sap[c] = df_sap[c].map(date_to_text_cell)

    df_sap["CRD_key"] = to_dt_series(df_sap["CRD"])
    df_sap["PD_key"]  = to_dt_series(df_sap["PD"])

    # ---------- MERGE ----------
    pick_cols = [
        "Order Status","Article No","Quantity",
        "Delay/Early - Confirmation CRD","Delay - PO PSDD Update","Delay - PO PD Update",
        "FPD","LPD","PSDD","PODD","CRD","PD",
        "Customer PO item","Line Aggregator","Shipment Method","Customer Number"
    ]
    inf_pick = df_infor_g[JOIN_KEYS + [c for c in pick_cols if c in df_infor_g.columns]].copy()
    inf_pick.rename(columns={c:f"infor {c}" for c in pick_cols if c in inf_pick.columns}, inplace=True)

    df = df_sap.merge(inf_pick, on=JOIN_KEYS, how="left")

    # ---------- DELAY MAPPING ----------
    for c in [
        "infor Delay/Early - Confirmation CRD",
        "infor Delay - PO PSDD Update",
        "infor Delay - PO PD Update"
    ]:
        if c in df.columns:
            df[c] = map_delay_series_to_code(df[c])

    # ---------- COMPARISON ----------
    df["Result_Quantity"] = equal_series(
        pd.to_numeric(df["Quanity"], errors="coerce"),
        pd.to_numeric(df["infor Quantity"], errors="coerce")
    )

    df["Result_Ship-to-Sort1"] = equal_series(
        df["Ship-to-Sort1"].astype(str),
        df["infor Customer Number"].astype(str)
    )

    # ---------- EXPORT (XLSXWRITER) ----------
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Data")
    out.seek(0)

    # ---------- STYLED EXPORT (OPENPYXL) ----------
    styled = None
    if EXCEL_STYLED_AVAILABLE:
        bio = io.BytesIO()
        with pd.ExcelWriter(bio, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Report")
            ws = writer.sheets["Report"]

            headers = list(ws.iter_rows(min_row=1, max_row=1))[0]
            for cell in headers:
                name = str(cell.value)
                if name.startswith("infor "):
                    cell.fill = PatternFill("solid", fgColor=INFOR_COLOR)
                elif name.startswith("Result"):
                    cell.fill = PatternFill("solid", fgColor=RESULT_COLOR)
                else:
                    cell.fill = PatternFill("solid", fgColor=OTHER_COLOR)
                cell.font = Font(bold=True, size=9)
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.font = Font(size=9)
                    cell.alignment = Alignment(horizontal="center", vertical="center")

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

        bio.seek(0)
        styled = bio

    return df, out, styled, out_name

# =========================
# STREAMLIT UI
# =========================
st.set_page_config("RSA - PGD Comparison", layout="wide")
st.title("RSA - PGD Comparison — FINAL (Styled)")

sap_file = st.file_uploader("Upload SAP file", ["xlsx","xls","csv"])
infor_files = st.file_uploader("Upload Infor file(s)", ["xlsx","xls","csv"], accept_multiple_files=True)

if sap_file and infor_files:
    sap_df = pd.read_excel(sap_file)
    infor_df = pd.concat([pd.read_excel(f) for f in infor_files], ignore_index=True)

    if st.button("Run comparison"):
        with st.spinner("Processing..."):
            df_final, raw_xlsx, styled_xlsx, fname = run_core_pipeline(sap_df, infor_df)

        st.success("Done")
        st.dataframe(df_final.head(200), use_container_width=True)

        st.download_button("⬇️ Download Excel", raw_xlsx.getvalue(), fname)
        if styled_xlsx:
            st.download_button("⬇️ Download Excel (Styled)", styled_xlsx.getvalue(), fname.replace(".xlsx","_styled.xlsx"))
