# app_streamlit_merged.py
# -*- coding: utf-8 -*-
"""
Merged Streamlit app:
- Core pipeline = script 1 (strict join on PO + CRD_key + PD_key) preserved.
- Adopts UI, multi-file Infor, PO Splitter, Dashboard, match_qty_nearest fallback, styled export from script 2.
- Export: primary xlsxwriter per-cell date (single-date -> Excel datetime), optional openpyxl styled export.
"""

import re
import io
import sys
import zipfile
from datetime import datetime, timedelta
from contextlib import nullcontext

import numpy as np
import pandas as pd
import streamlit as st

# timezone
try:
    from zoneinfo import ZoneInfo
    _tz = ZoneInfo("Asia/Jakarta")
except Exception:
    _tz = None

# Optional openpyxl styling
EXCEL_STYLED_AVAILABLE = True
try:
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except Exception:
    EXCEL_STYLED_AVAILABLE = False

# =====================
# Config / constants
# =====================
JOIN_KEYS = ["PO No.(Full)", "CRD_key", "PD_key"]
DATE_COLS_INFOR = ["Issue Date","FPD","LPD","PSDD","PODD","CRD","PD"]
DATE_COLS_SAP   = ["Document Date","FPD","LPD","PSDD","PODD","FCR Date","PO Date","CRD","PD","Actual PGI"]

# desired order adapted from script 1 (kept names)
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
    "MDP","PDP","SDP","Article Lead time","Ship-to-Sort1","Ship-to Country","Ship to Name","infor Shipment Method",
    "Document Date","FPD","infor FPD","Result FPD","LPD","infor LPD","Result LPD",
    "CRD","infor CRD","Result CRD","PSDD","infor PSDD","Result PSDD",
    "PODD","infor PODD","Result PODD","FCR Date","PD","infor PD","Result PD",
    "PO Date","Actual PGI","Segment","S&P LPD","Currency"
]

# for styled export (colors)
INFOR_COLOR  = "FFF9F16D"
RESULT_COLOR = "FFC6EFCE"
OTHER_COLOR  = "FFD9D9D9"
DATE_FMT_OPENPYXL = "m/d/yyyy"

# blanks
BLANKS = {"(blank)", "blank", "", "--", " -- ", " --"}

# =====================
# Helpers from script1 (kept)
# =====================
def to_dt_series(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce").dt.normalize()

def to_dt_scalar(x):
    ts = pd.to_datetime(x, errors="coerce")
    if pd.isna(ts): return pd.NaT
    return pd.to_datetime(ts).normalize()

def fmt_dt(dt: pd.Timestamp) -> str:
    if pd.isna(dt): return ""
    return dt.strftime("%m/%d/%Y")

def date_concat(series: pd.Series) -> str:
    dts = to_dt_series(series).dropna()
    if dts.empty:
        return np.nan
    uniq = sorted(set(dts))
    return " | ".join(fmt_dt(x) for x in uniq)

def date_to_text_cell(val) -> str:
    return fmt_dt(to_dt_scalar(val))

def sum_keep_nan(s: pd.Series):
    return s.sum(min_count=1)

def keep_or_join(s: pd.Series):
    vals = pd.unique(s.dropna())
    if len(vals) == 0: return np.nan
    if len(vals) == 1: return vals[0]
    return " | ".join(map(str, vals))

def extract_digits(x):
    if pd.isna(x): return np.nan
    s = str(x).strip()
    if s.lower() in {b.lower() for b in BLANKS}: return np.nan
    m = re.search(r"\d+", s)
    return m.group(0) if m else np.nan

def split_date_set(x):
    if pd.isna(x):
        return set()
    parts = [p.strip() for p in str(x).split("|")]
    out = set()
    for p in parts:
        if not p:
            continue
        dt = pd.to_datetime(p, errors="coerce")
        if not pd.isna(dt):
            out.add(pd.to_datetime(dt).normalize())
    return out

def to_excel_col(n):
    s = ""; n += 1
    while n:
        n, r = divmod(n-1, 26)
        s = chr(r + 65) + s
    return s

# =====================
# Utility functions (adopted from script2)
# =====================
def today_str_id():
    return (datetime.utcnow() + timedelta(hours=7)).strftime("%Y%m%d")

def parse_input(text: str, split_mode: str = "auto"):
    text = text.strip()
    if not text: return []
    if split_mode == "newline": raw = text.splitlines()
    elif split_mode == "comma": raw = text.split(",")
    elif split_mode == "semicolon": raw = text.split(";")
    elif split_mode == "whitespace": raw = re.split(r"\s+", text)
    else:
        if "\n" in text:
            raw = re.split(r"[\r\n]+", text)
            split_more = []
            for line in raw:
                line = line.strip()
                if not line: continue
                if ("," in line) or (";" in line):
                    split_more.extend(re.split(r"[,;]", line))
                else:
                    split_more.append(line)
            raw = split_more
        elif ("," in text) or (";" in text):
            raw = re.split(r"[,;]", text)
        else:
            raw = re.split(r"\s+", text)
    return [x.strip() for x in raw if str(x).strip() != ""]

def normalize_items(items, keep_only_digits=False, upper_case=False, strip_prefix_suffix=False):
    normed = []
    for it in items:
        s = str(it)
        if strip_prefix_suffix: s = re.sub(r"^\W+|\W+$", "", s)
        if keep_only_digits: s = re.sub(r"\D+", "", s)
        if upper_case: s = s.upper()
        s = s.strip()
        if s != "": normed.append(s)
    return normed

def chunk_list(items, size):
    return [items[i:i+size] for i in range(0, len(items), size)]

def df_from_list(items, col_name="PO"):
    return pd.DataFrame({col_name: items})

def make_zip_bytes(chunks, basename="chunk", as_csv=True, col_name="PO"):
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for idx, part in enumerate(chunks, start=1):
            if as_csv:
                df = df_from_list(part, col_name=col_name)
                csv_bytes = df.to_csv(index=False).encode("utf-8")
                zf.writestr(f"{basename}_{idx:02d}.csv", csv_bytes)
            else:
                buf = io.StringIO()
                for ln in part: buf.write(f"{ln}\n")
                zf.writestr(f"{basename}_{idx:02d}.txt", buf.getvalue().encode("utf-8"))
    mem.seek(0);  return mem

# =====================
# Loading helpers
# =====================
@st.cache_data(show_spinner=False)
def read_excel_file_bytes(uploaded):
    # uploaded can be file-like or path
    if hasattr(uploaded, "read"):
        uploaded.seek(0)
        return pd.read_excel(uploaded, engine="openpyxl")
    else:
        return pd.read_excel(str(uploaded), engine="openpyxl")

@st.cache_data(show_spinner=False)
def read_csv_file_bytes(uploaded):
    # try encodings
    if hasattr(uploaded, "read"):
        for enc in ("utf-8", "utf-8-sig", "latin1"):
            try:
                uploaded.seek(0)
                return pd.read_csv(uploaded, encoding=enc)
            except Exception:
                continue
        uploaded.seek(0)
        return pd.read_csv(uploaded)
    else:
        return pd.read_csv(str(uploaded))

# =====================
# match_qty_nearest (from script2) - used as fallback / option
# =====================
def match_qty_nearest(df_sap, df_infor, key="PO No.(Full)", qty_col="Quanity", infor_qty_col="infor Quantity"):
    """
    - If SAP has 1 row and Infor multiple -> choose Infor row with closest quantity and blank SAP quantity on other rows.
    - Keep original columns where possible. Works with SAP Quanity typo.
    """
    df_sap = df_sap.copy()
    df_sap["___sap_row_id"] = np.arange(len(df_sap))
    merged = df_sap.merge(df_infor, how="left", left_on=key, right_on="PO No.(Full)" if "PO No.(Full)" in df_infor.columns else key)
    if merged.empty:
        return merged
    out_rows = []
    for po, group in merged.groupby(key, sort=False):
        n_sap_rows = group["___sap_row_id"].nunique() if "___sap_row_id" in group.columns else 1
        if n_sap_rows > 1:
            out_rows.extend(group.drop(columns=["___sap_row_id"], errors="ignore").to_dict("records"))
            continue
        if len(group) == 1:
            out_rows.append(group.drop(columns=["___sap_row_id"], errors="ignore").iloc[0].to_dict())
            continue
        try:
            sap_qty = pd.to_numeric(group.iloc[0].get(qty_col, 0), errors="coerce")
        except Exception:
            sap_qty = pd.to_numeric(group.iloc[0].get("Quantity", 0), errors="coerce")
        diffs = (pd.to_numeric(group.get(infor_qty_col, 0), errors="coerce") - sap_qty).abs()
        idx_min = diffs.idxmin() if diffs.notna().any() else group.index[0]
        for idx, row in group.iterrows():
            row = row.copy()
            if idx != idx_min:
                # blank out SAP quantity columns to avoid double-count
                if qty_col in row.index:
                    row[qty_col] = np.nan
                if "Quantity" in row.index:
                    row["Quantity"] = np.nan
            out_rows.append(row.drop(labels=["___sap_row_id"], errors="ignore").to_dict())
    return pd.DataFrame(out_rows)

# =====================
# Core pipeline adapted from script1 but with options
# =====================
def run_core_pipeline(df_sap_raw, df_infor_raw_all, *,
                     prefer_strict_join=True,
                     fallback_match_qty=True,
                     export_filename_prefix=None):
    """
    - df_sap_raw: DataFrame loaded from SAP input
    - df_infor_raw_all: DataFrame combined from all Infor files (un-grouped)
    - prefer_strict_join: if True, try strict join on JOIN_KEYS first
    - fallback_match_qty: if strict join yields poor matches, allow fallback to match_qty_nearest joining by PO only
    Returns df_final (DataFrame) and bytes_io (xlsx with xlsxwriter per-cell dates) and out_filename
    """
    # copy inputs
    df_sap = df_sap_raw.copy()
    df_infor = df_infor_raw_all.copy()

    # preserve original OUT name
    _today = datetime.now(_tz) if _tz else datetime.now()
    _today_str = _today.strftime("%Y%m%d")
    OUT_JOINED = export_filename_prefix or f"RSA - PGD Comparison Tracking Report - {_today_str}.xlsx"

    # ======= Normalisasi Infor (as script1) =======
    rename_cols = {
        'Order #': 'PO No.(Full)',
        'Line Aggregator': 'Customer PO item',
        'Article Number': 'Article No',
        'Country/Region': 'Ship-to Country',
        'Customer Request Date (CRD)': 'CRD',
        'Plan Date': 'PD',
        'PO Statistical Delivery Date (PSDD)': 'PSDD',
        'First Production Date': 'FPD',
        'Last Production Date': 'LPD',
        'Production Lead Time': 'Lead time',
        'Class Code': 'Classification Code',
        'Delay - Confirmation': 'Delay/Early - Confirmation CRD',
        'Delay - PO Del Update': 'Delay - PO PSDD Update',
        'Grand Total': 'Quantity',
        'Delivery Delay Pd': 'Delay - PO PD Update',
        'Shipment Method': 'Shipment Method',
    }
    df_infor = df_infor.rename(columns=rename_cols)

    # standard PD confirmation rename if present
    if "Confirmation Delay Pd" in df_infor.columns and "Delay/Early - Confirmation PD" not in df_infor.columns:
        df_infor = df_infor.rename(columns={"Confirmation Delay Pd": "Delay/Early - Confirmation PD"})

    # drop Quantity == 0
    if "Quantity" in df_infor.columns:
        df_infor = df_infor[df_infor["Quantity"].fillna(0) != 0].copy()

    # detect size columns like script1
    size_pat  = re.compile(r'^(?:[1-9]|1[0-8])(?:K|-K|-)?$')
    size_cols = [c for c in df_infor.columns if size_pat.match(str(c))]
    sum_cols   = size_cols + (["Quantity"] if "Quantity" in df_infor.columns else [])
    other_cols = [c for c in df_infor.columns if c not in (["Issue Date","PO No.(Full)","Model Name","Article No","Ship-to Country","CRD","PD"] + sum_cols)]

    # aggregator
    agg_infor = {}
    for col in sum_cols:
        agg_infor[col] = sum_keep_nan
    for col in other_cols:
        if col in DATE_COLS_INFOR:
            agg_infor[col] = date_concat
        else:
            agg_infor[col] = keep_or_join

    meta_cols = ["Issue Date", "PO No.(Full)", "Model Name", "Article No", "Ship-to Country", "CRD", "PD"]
    missing = [c for c in meta_cols if c not in df_infor.columns]
    if missing:
        # If Infor not having Issue Date or some meta but has PO No., fallback to group by PO No.
        if "PO No.(Full)" in df_infor.columns:
            # fallback grouping by PO only (like script2 behavior)
            grp_cols = ["PO No.(Full)"]
            # pick sane defaults: sum quantity, first for others
            df_infor_grouped = df_infor.groupby(grp_cols, dropna=False).agg({
                'Quantity':'sum'
            }).reset_index()
            # copy other important columns via keep_or_join
            for c in ["Model Name","Article No","Country/Region","Customer Request Date (CRD)","Plan Date"]:
                if c in df_infor.columns:
                    df_infor_grouped[c] = df_infor.groupby("PO No.(Full)")[c].agg(keep_or_join).values
            # rename to expected names
            if "Customer Request Date (CRD)" in df_infor_grouped.columns:
                df_infor_grouped = df_infor_grouped.rename(columns={"Customer Request Date (CRD)": "CRD"})
            if "Plan Date" in df_infor_grouped.columns:
                df_infor_grouped = df_infor_grouped.rename(columns={"Plan Date": "PD"})
        else:
            raise ValueError(f"Kolom meta tidak ditemukan di Infor: {missing}")
    else:
        # normal grouping
        df_inf_num = df_infor.copy()
        if sum_cols:
            df_inf_num[sum_cols] = df_inf_num[sum_cols].apply(pd.to_numeric, errors="coerce")
        df_infor_grouped = df_inf_num.groupby(meta_cols, dropna=False).agg(agg_infor).reset_index()
        if "Customer PO item" in df_infor_grouped.columns and "Line Aggregator" not in df_infor_grouped.columns:
            df_infor_grouped["Line Aggregator"] = df_infor_grouped["Customer PO item"]

    # make CRD_key/PD_key
    if "CRD" in df_infor_grouped.columns:
        df_infor_grouped["CRD_key"] = to_dt_series(df_infor_grouped["CRD"])
    else:
        df_infor_grouped["CRD_key"] = pd.NaT
    if "PD" in df_infor_grouped.columns:
        df_infor_grouped["PD_key"] = to_dt_series(df_infor_grouped["PD"])
    else:
        df_infor_grouped["PD_key"] = pd.NaT

    # ======= SAP row-level (date display as text) =======
    df_sap = df_sap.copy()
    for col in DATE_COLS_SAP:
        if col in df_sap.columns:
            df_sap[col] = df_sap[col].map(date_to_text_cell)
    # create keys
    df_sap["CRD_key"] = to_dt_series(df_sap["CRD"]) if "CRD" in df_sap.columns else pd.NaT
    df_sap["PD_key"]  = to_dt_series(df_sap["PD"]) if "PD" in df_sap.columns else pd.NaT

    # ======= prepare infor pick with prefix 'infor ' as script1 did =======
    infor_cols_for_merge = [
        "Order Status","Article No","LPD","PODD","PSDD","FPD","CRD","PD",
        "Delay/Early - Confirmation CRD","Delay - PO PSDD Update","Delay - PO PD Update",
        "Quantity","Shipment Method","Issue Date",
        "Customer PO item","Line Aggregator"
    ]
    if "Delay/Early - Confirmation PD" in df_infor_grouped.columns:
        infor_cols_for_merge = ["Delay/Early - Confirmation PD"] + infor_cols_for_merge

    inf_pick_cols = [c for c in infor_cols_for_merge if c in df_infor_grouped.columns]
    inf_pick = df_infor_grouped[["PO No.(Full)","CRD_key","PD_key"] + inf_pick_cols].copy()
    pref_map = {c: f"infor {c}" for c in inf_pick_cols}
    inf_pick = inf_pick.rename(columns=pref_map)

    # ----- Try strict join first if desired -----
    df_join = None
    if prefer_strict_join:
        try:
            df_join = df_sap.merge(inf_pick, on=JOIN_KEYS, how="left")
        except Exception:
            df_join = df_sap.merge(inf_pick, on=["PO No.(Full)"], how="left")  # fallback

    # If strict join was not done or yields empty matches, optionally fallback to tolerant join using match_qty_nearest
    if df_join is None or df_join.shape[0] == 0 or df_join["infor Quantity"].isna().all():
        if fallback_match_qty:
            # Prepare infor for match_qty: ensure infor has 'PO No.(Full)'
            df_infor_for_match = df_infor_grouped.copy()
            # ensure infor has column names for matching function
            if "PO No.(Full)" not in df_infor_for_match.columns:
                # try to rename if exists 'Order #' originally
                if "Order #" in df_infor.columns:
                    df_infor_for_match = df_infor_for_match.rename(columns={"Order #": "PO No.(Full)"})
            # create 'infor Quantity' column name expected by match_qty_nearest
            if "Quantity" in df_infor_for_match.columns and "infor Quantity" not in df_infor_for_match.columns:
                df_infor_for_match = df_infor_for_match.rename(columns={"Quantity": "infor Quantity"})
            # ensure SAP has Quanity column name (script1 uses 'Quanity')
            if "Quanity" not in df_sap.columns and "Quantity" in df_sap.columns:
                df_sap = df_sap.rename(columns={"Quantity": "Quanity"})
            # run tolerant matching
            df_join = match_qty_nearest(df_sap, df_infor_for_match, key="PO No.(Full)", qty_col="Quanity", infor_qty_col="infor Quantity")
            # If match_qty_nearest returns dataframe with PO-level columns (no prefix), add 'infor ' prefix where needed
            # Keep other columns as-is.
        else:
            # final fallback: simple left merge on PO No.(Full)
            df_join = df_sap.merge(inf_pick, on=["PO No.(Full)"], how="left")

    # If df_join is still None for any reason, raise
    if df_join is None:
        raise RuntimeError("Merge failed — no join result produced.")

    # ======= Mapping delay code (script1 logic) =======
    code_mapping = {
        '161':'01-0161','84':'03-0084','68':'02-0068','64':'04-0064','62':'02-0062','61':'01-0061',
        '51':'03-0051','46':'03-0046','7':'02-0007','3':'03-0003','2':'01-0002','1':'01-0001',
        '4':'04-0004','8':'02-0008','10':'04-0010','49':'03-0049','90':'04-0090','63':'03-0063',
    }
    def map_delay_series_to_code(s: pd.Series, code_mapping: dict) -> pd.Series:
        base = s.apply(extract_digits)
        mapped = base.map(code_mapping)
        return mapped.where(mapped.notna(), base)

    for col in ["infor Delay/Early - Confirmation PD",
                "infor Delay/Early - Confirmation CRD",
                "infor Delay - PO PSDD Update",
                "infor Delay - PO PD Update"]:
        if col in df_join.columns:
            mapped_col = f"{col} (Mapped)"
            df_join[mapped_col] = map_delay_series_to_code(df_join[col], code_mapping)
            df_join[col] = df_join[mapped_col]
            df_join.drop(columns=[mapped_col], inplace=True)

    # ======= Compare logic (script1) =======
    def norm_num(s): return pd.to_numeric(s, errors="coerce")
    def norm_str(s):
        s = s.astype(str).str.strip()
        s = s.replace(list(BLANKS), np.nan)
        return s
    def norm_delay(s): return s.apply(extract_digits)
    def equal_series(a,b): return a.eq(b) | (a.isna() & b.isna())
    def compare_dates_as_sets(s_left, s_right):
        left_sets  = s_left.apply(split_date_set)
        right_sets = s_right.apply(split_date_set)
        return (left_sets == right_sets).fillna(False)

    df = df_join.copy()

    # Quanity vs infor Quantity -> Result_Quantity
    if "Quanity" in df.columns and "infor Quantity" in df.columns:
        df["Result_Quantity"] = equal_series(norm_num(df["Quanity"]), norm_num(df["infor Quantity"])).fillna(False)

    # Article No
    if "Article No" in df.columns and "infor Article No" in df.columns:
        df["Result Article No"] = equal_series(norm_str(df["Article No"]), norm_str(df["infor Article No"])).fillna(False)

    # Delay fields
    if "Delay/Early - Confirmation PD" in df.columns and "infor Delay/Early - Confirmation PD" in df.columns:
        df["Result Delay/Early - Confirmation PD"] = equal_series(
            norm_delay(df["Delay/Early - Confirmation PD"]), norm_delay(df["infor Delay/Early - Confirmation PD"])
        ).fillna(False)

    if "Delay/Early - Confirmation CRD" in df.columns and "infor Delay/Early - Confirmation CRD" in df.columns:
        df["Result Delay/Early - Confirmation CRD"] = equal_series(
            norm_delay(df["Delay/Early - Confirmation CRD"]), norm_delay(df["infor Delay/Early - Confirmation CRD"])
        ).fillna(False)

    if "Delay - PO PSDD Update" in df.columns and "infor Delay - PO PSDD Update" in df.columns:
        df["Result Delay - PO PSDD Update"] = equal_series(
            norm_delay(df["Delay - PO PSDD Update"]), norm_delay(df["infor Delay - PO PSDD Update"])
        ).fillna(False)

    if "Delay - PO PD Update" in df.columns and "infor Delay - PO PD Update" in df.columns:
        df["Result Delay - PO PD Update"] = equal_series(
            norm_delay(df["Delay - PO PD Update"]), norm_delay(df["infor Delay - PO PD Update"])
        ).fillna(False)

    # date comparisons (set-equality)
    date_pairs = [
        ("FPD","infor FPD","Result FPD"),
        ("LPD","infor LPD","Result LPD"),
        ("CRD","infor CRD","Result CRD"),
        ("PSDD","infor PSDD","Result PSDD"),
        ("PODD","infor PODD","Result PODD"),
        ("PD","infor PD","Result PD"),
    ]
    for left, right, outcol in date_pairs:
        if left in df.columns and right in df.columns:
            df[outcol] = compare_dates_as_sets(df[left], df[right])

    # Shipment Method compare
    if "Shipment Method" in df.columns and "infor Shipment Method" in df.columns:
        df["Result Shipment Method"] = equal_series(norm_str(df["Shipment Method"]), norm_str(df["infor Shipment Method"])).fillna(False)

    # ensure Line Aggregator present
    if "Line Aggregator" not in df.columns and "infor Line Aggregator" in df.columns:
        df["Line Aggregator"] = df["infor Line Aggregator"]

    # reorder
    present = [c for c in DESIRED_ORDER if c in df.columns]
    rest     = [c for c in df.columns if c not in present]
    df_final = df[present + rest]

    # ======= Export to xlsxwriter with per-cell date handling (script1) =======
    DATE_FMT = "mm/dd/yyyy"
    date_display_cols = [
        "Document Date","FPD","LPD","PSDD","PODD","FCR Date","PD","PO Date","Actual PGI",
        "infor FPD","infor LPD","infor CRD","infor PSDD","infor PODD","infor PD",
        "CRD"
    ]
    date_display_cols = [c for c in date_display_cols if c in df_final.columns]

    def is_single_date_text(s: str) -> bool:
        if not isinstance(s, str): return False
        if " | " in s: return False
        dt = pd.to_datetime(s, errors="coerce")
        return not pd.isna(dt)

    out = io.BytesIO()
    try:
        import xlsxwriter
        with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
            df_final.to_excel(writer, index=False, sheet_name="Data", startrow=0, startcol=0, header=True)
            wb  = writer.book
            ws  = writer.sheets["Data"]

            fmt_bool_T = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
            fmt_bool_F = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
            fmt_date   = wb.add_format({"num_format": DATE_FMT})

            nrows, ncols = df_final.shape
            # rewrite date columns cell-by-cell
            for cidx, col in enumerate(df_final.columns):
                if col in date_display_cols:
                    for ridx, val in enumerate(df_final[col].tolist(), start=1):
                        if pd.isna(val):
                            ws.write_string(ridx, cidx, "")
                        else:
                            s = str(val).strip()
                            if is_single_date_text(s):
                                dt = pd.to_datetime(s, errors="coerce")
                                if not pd.isna(dt):
                                    ws.write_datetime(ridx, cidx, dt.to_pydatetime(), fmt_date)
                                else:
                                    ws.write_string(ridx, cidx, s)
                            else:
                                ws.write_string(ridx, cidx, s)

            # conditional formatting for Result columns (boolean True/False)
            res_cols = [c for c in df_final.columns if c.startswith("Result ")]
            for col in res_cols:
                cidx = df_final.columns.get_loc(col)
                rng  = f"{to_excel_col(cidx)}2:{to_excel_col(cidx)}{nrows+1}"
                # xlsxwriter conditional_format requires comparing to string 'TRUE'/'FALSE' or boolean values.
                ws.conditional_format(rng, {"type":"cell","criteria":"==","value":"TRUE","format":fmt_bool_T})
                ws.conditional_format(rng, {"type":"cell","criteria":"==","value":"FALSE","format":fmt_bool_F})

            # set column widths
            for idx, col in enumerate(df_final.columns, start=1):
                if col.startswith("Result "):
                    ws.set_column(idx-1, idx-1, 16)
                elif col in date_display_cols:
                    ws.set_column(idx-1, idx-1, 12)
                else:
                    ws.set_column(idx-1, idx-1, 18)

            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, nrows, ncols-1)
            writer.save()
        out.seek(0)
    except Exception as e:
        # fallback openpyxl simple write
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            df_final.to_excel(writer, index=False, sheet_name="Data")
        out.seek(0)

    # Also prepare styled openpyxl export (optional) if available
    styled_bytes = None
    if EXCEL_STYLED_AVAILABLE:
        try:
            bio = io.BytesIO()
            with pd.ExcelWriter(bio, engine="openpyxl") as writer:
                df_final.to_excel(writer, index=False, sheet_name="Report")
                ws = writer.sheets["Report"]
                # header styling
                header_cells = list(ws.iter_rows(min_row=1, max_row=1, values_only=False))[0]
                idx_by_name = {c.value: i+1 for i, c in enumerate(header_cells)}
                for cell in header_cells:
                    col_name = str(cell.value)
                    if col_name.startswith("infor ") or col_name.startswith("Infor "):
                        fill = PatternFill("solid", fgColor=INFOR_COLOR)
                    elif col_name.startswith("Result ") or col_name.startswith("Result_"):
                        fill = PatternFill("solid", fgColor=RESULT_COLOR)
                    else:
                        fill = PatternFill("solid", fgColor=OTHER_COLOR)
                    cell.fill = fill
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = Font(name="Calibri", size=9, bold=True)
                # row styles
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                    for cell in row:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.font = Font(name="Calibri", size=9)
                # date formats
                for date_col in date_display_cols:
                    if date_col in idx_by_name:
                        cidx = idx_by_name[date_col]
                        for r in range(2, ws.max_row + 1):
                            cell = ws.cell(row=r, column=cidx)
                            if cell.value not in ("", None):
                                try:
                                    # if string date, try parse
                                    if isinstance(cell.value, str):
                                        dt = pd.to_datetime(cell.value, errors="coerce")
                                        if not pd.isna(dt):
                                            cell.value = dt
                                    cell.number_format = DATE_FMT_OPENPYXL
                                except Exception:
                                    pass
                # auto width
                for col_idx in range(1, ws.max_column + 1):
                    col_letter = get_column_letter(col_idx)
                    maxlen = 0
                    for cell in ws[col_letter]:
                        v = "" if cell.value is None else str(cell.value)
                        maxlen = max(maxlen, len(v))
                    ws.column_dimensions[col_letter].width = min(max(9, maxlen + 2), 60)
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
            bio.seek(0)
            styled_bytes = bio
        except Exception:
            styled_bytes = None

    return df_final, out, styled_bytes, OUT_JOINED

# =====================
# Streamlit UI
# =====================
st.set_page_config(page_title="RSA - PGD Comparison (Merged)", layout="wide")
st.title("RSA - PGD Comparison — Merged (Script1 core + Script2 UX & features)")

with st.sidebar:
    st.header("Upload files")
    sap_file = st.file_uploader("Upload SAP Excel (.xlsx/.xls/.csv)", type=["xlsx","xls","csv"], key="sap_file")
    infor_files = st.file_uploader("Upload Infor files (CSV/XLSX) - can upload multiple", type=["csv","xlsx","xls"], accept_multiple_files=True, key="infor_files")
    st.markdown("---")
    st.header("Options")
    prefer_strict = st.checkbox("Prefer strict join (PO + CRD_key + PD_key)", value=True)
    fallback_match = st.checkbox("Enable fallback tolerant matching (match_qty_nearest)", value=True)
    export_style_choice = st.selectbox("Export Styled Excel (if available)", options=["xlsxwriter per-cell date (default)","openpyxl styled export (header colors)"], index=0)
    st.markdown("Tips: gunakan strict join bila data tanggal konsisten; gunakan fallback jika Infor memiliki multi-row per PO.")
    st.markdown("---")
    st.header("PO Splitter")
    show_po_splitter = st.checkbox("Enable PO Splitter tab", value=True)

# Tabs: Comparison / PO Splitter / Dashboard
tab1, tab2, tab3 = st.tabs(["📦 PGD Comparison", "🧩 PO Splitter", "📈 Sales Analytics (SAP)"])

# ---------------- Tab1: Comparison ----------------
with tab1:
    st.markdown("### 1) Upload & Run Pipeline")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("SAP preview")
        df_sap_preview = None
        if sap_file:
            try:
                if sap_file.name.lower().endswith((".xls",".xlsx")):
                    df_sap_preview = read_excel_file_bytes(sap_file)
                else:
                    df_sap_preview = read_csv_file_bytes(sap_file)
                st.dataframe(df_sap_preview.head(200))
            except Exception as e:
                st.error("Gagal membaca SAP file.")
                st.exception(e)
        else:
            st.info("Unggah SAP file di sidebar.")

    with col2:
        st.subheader("Infor preview (combined)")
        df_infor_preview = None
        if infor_files:
            try:
                # read multiple infor files and concat
                df_list = []
                for f in infor_files:
                    if f.name.lower().endswith((".xls", ".xlsx")):
                        df_list.append(read_excel_file_bytes(f))
                    else:
                        df_list.append(read_csv_file_bytes(f))
                df_infor_preview = pd.concat(df_list, ignore_index=True)
                st.dataframe(df_infor_preview.head(200))
            except Exception as e:
                st.error("Gagal membaca Infor files.")
                st.exception(e)
        else:
            st.info("Unggah minimal 1 Infor file di sidebar (bisa multiple CSV/XLSX).")

    run_button = st.button("Run pipeline", type="primary")
    if run_button:
        if sap_file is None or not infor_files:
            st.error("Harap upload SAP file dan minimal 1 Infor file.")
        else:
            try:
                # load SAP & Infor combined
                if sap_file.name.lower().endswith((".xls",".xlsx")):
                    df_sap_load = read_excel_file_bytes(sap_file)
                else:
                    df_sap_load = read_csv_file_bytes(sap_file)
                # combine infor files
                infor_dfs = []
                for f in infor_files:
                    if f.name.lower().endswith((".xls",".xlsx")):
                        infor_dfs.append(read_excel_file_bytes(f))
                    else:
                        infor_dfs.append(read_csv_file_bytes(f))
                df_infor_all = pd.concat(infor_dfs, ignore_index=True)

                with st.spinner("Running pipeline..."):
                    df_final, bytes_xlsx, styled_bytes, out_name = run_core_pipeline(
                        df_sap_load, df_infor_all,
                        prefer_strict_join=prefer_strict,
                        fallback_match_qty=fallback_match,
                        export_filename_prefix=None
                    )
                st.success("Pipeline selesai.")
                st.markdown("#### Preview result (top 200 rows)")
                st.dataframe(df_final.head(200), use_container_width=True)

                # download buttons
                st.download_button("⬇️ Download Excel (primary - xlsxwriter per-cell)", data=bytes_xlsx.getvalue(),
                                   file_name=out_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                if EXCEL_STYLED_AVAILABLE and styled_bytes is not None:
                    st.download_button("⬇️ Download Excel (styled - openpyxl)", data=styled_bytes.getvalue(),
                                       file_name=out_name.replace(".xlsx","_styled.xlsx"),
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                csv_name = out_name.replace(".xlsx", ".csv")
                st.download_button("⬇️ Download CSV (basic)", data=df_final.to_csv(index=False).encode("utf-8"), file_name=csv_name, mime="text/csv")

                # quick summary of Result* columns
                st.markdown("### Summary Result*")
                res_bool_cols = [c for c in df_final.columns if c.startswith("Result ")]
                if res_bool_cols:
                    summary = {}
                    for c in res_bool_cols:
                        col = df_final[c]
                        true_count = int(col.eq(True).sum()) if col.dtype == bool else int(col.eq("TRUE").sum()) if col.dtype == object else int(col.eq(True).sum())
                        false_count = int(col.eq(False).sum()) if col.dtype == bool else int(col.eq("FALSE").sum()) if col.dtype == object else int(col.eq(False).sum())
                        summary[c] = {"TRUE": true_count, "FALSE": false_count}
                    st.json(summary)
                else:
                    st.info("Tidak ditemukan kolom Result* di hasil.")

            except Exception as e:
                st.error("Pipeline gagal.")
                st.exception(e)

# ---------------- Tab2: PO Splitter ----------------
with tab2:
    st.markdown("### PO Splitter — bagi daftar PO ke beberapa file/zip")
    st.caption("Tempel daftar PO, app akan membagi ke chunks (default 5000).")
    split_mode = st.selectbox("Mode pemisah", ["auto", "newline", "comma", "semicolon", "whitespace"], index=0)
    chunk_size = st.number_input("Maks PO per bagian", min_value=1, max_value=1_000_000, value=5000, step=1)
    drop_duplicates = st.checkbox("Hapus duplikat (preserve first)", value=False)
    keep_only_digits = st.checkbox("Keep only digits", value=False)
    upper_case = st.checkbox("Upper-case", value=False)
    strip_prefix_suffix = st.checkbox("Strip prefix/suffix non-alphanum", value=False)

    input_text = st.text_area("Tempel daftar PO di sini:", height=220, placeholder="PO001\nPO002\n... atau comma/space separated")
    if st.button("Proses & Bagi PO"):
        items = parse_input(input_text, split_mode=split_mode)
        original_count = len(items)
        if keep_only_digits or upper_case or strip_prefix_suffix:
            items = normalize_items(items, keep_only_digits=keep_only_digits, upper_case=upper_case, strip_prefix_suffix=strip_prefix_suffix)
        if drop_duplicates:
            items = list(dict.fromkeys(items))
        total = len(items)
        st.write(f"Total input: {original_count} → setelah proses: {total}")
        if total == 0:
            st.warning("Tidak ada PO.")
        else:
            parts = chunk_list(items, int(chunk_size))
            st.success(f"Terbagi menjadi {len(parts)} bagian.")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            zip_csv = make_zip_bytes(parts, basename=f"PO_chunk_{timestamp}", as_csv=True)
            st.download_button("Unduh ZIP (CSV semua parts)", data=zip_csv, file_name=f"PO_chunks_csv_{timestamp}.zip", mime="application/zip")
            zip_txt = make_zip_bytes(parts, basename=f"PO_chunk_{timestamp}", as_csv=False)
            st.download_button("Unduh ZIP (TXT semua parts)", data=zip_txt, file_name=f"PO_chunks_txt_{timestamp}.zip", mime="application/zip")
            st.markdown("#### Pratinjau per bagian")
            for idx, part in enumerate(parts, start=1):
                with st.expander(f"Part {idx} — {len(part)} PO"):
                    dfp = df_from_list(part, col_name="PO")
                    st.dataframe(dfp, use_container_width=True)
                    st.download_button(f"Unduh Part {idx} CSV", data=dfp.to_csv(index=False).encode("utf-8"), file_name=f"PO_chunk_{idx:02d}.csv", mime="text/csv")

# ---------------- Tab3: Sales Analytics (SAP-only)
with tab3:
    st.markdown("### Sales Analytics Dashboard (SAP-only)")
    sap_dash = st.file_uploader("Upload SAP Excel (.xlsx) for dashboard", type=["xlsx","xls","csv"], key="sap_dash")
    if sap_dash:
        try:
            if sap_dash.name.lower().endswith((".xls",".xlsx")):
                sap_df = read_excel_file_bytes(sap_dash)
            else:
                sap_df = read_csv_file_bytes(sap_dash)
            # normalize: Quanity -> Quantity
            if "Quanity" in sap_df.columns and "Quantity" not in sap_df.columns:
                sap_df = sap_df.rename(columns={"Quanity":"Quantity"})
            # convert date-like columns
            for col in DATE_COLS_SAP:
                if col in sap_df.columns:
                    sap_df[col] = pd.to_datetime(sap_df[col], errors="coerce")
            # sidebar filters
            with st.sidebar:
                st.header("Dashboard Filters")
                date_axis_col = st.selectbox("Date axis", options=[c for c in ["CRD","LPD","PD","PO Date","Document Date","PSDD","FPD","PODD","Actual PGI"] if c in sap_df.columns], index=0)
                min_d = sap_df[date_axis_col].min()
                max_d = sap_df[date_axis_col].max()
                sd, ed = st.date_input("Date range", value=(min_d.date() if pd.notna(min_d) else datetime(2025,1,1).date(), max_d.date() if pd.notna(max_d) else datetime(2025,12,31).date()))
                def options(col):
                    return sorted([x for x in sap_df[col].dropna().astype(str).unique()]) if col in sap_df.columns else []
                segs = st.multiselect("Segment", options("Segment"))
                sites = st.multiselect("Site", options("Site"))
                brands = st.multiselect("Brand FTY Name", options("Brand FTY Name"))
            # filter df
            df_dash = sap_df.copy()
            df_dash[date_axis_col] = pd.to_datetime(df_dash[date_axis_col], errors="coerce")
            if sd and ed:
                df_dash = df_dash[(df_dash[date_axis_col] >= pd.Timestamp(sd)) & (df_dash[date_axis_col] <= pd.Timestamp(ed) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1))]
            def apply_multi(df_in, col, vals):
                if vals and col in df_in.columns:
                    return df_in[df_in[col].astype(str).isin(vals)]
                return df_in
            for col, vals in [("Segment", segs), ("Site", sites), ("Brand FTY Name", brands)]:
                df_dash = apply_multi(df_dash, col, vals)
            # metrics
            df_dash["__Value__"] = (pd.to_numeric(df_dash.get("Unit Price", 0), errors="coerce").fillna(0.0) *
                                    pd.to_numeric(df_dash.get("Quantity", 0), errors="coerce").fillna(0.0))
            total_orders = df_dash["PO No.(Full)"].nunique() if "PO No.(Full)" in df_dash.columns else len(df_dash)
            total_qty = int(pd.to_numeric(df_dash.get("Quantity", 0), errors="coerce").fillna(0).sum())
            total_value = float(df_dash["__Value__"].sum())
            avg_up = float(pd.to_numeric(df_dash.get("Unit Price", 0), errors="coerce").fillna(0).mean())
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Unique PO", f"{total_orders:,}")
            c2.metric("Total Quantity", f"{total_qty:,}")
            c3.metric("Total Value", f"${total_value:,.0f}")
            c4.metric("Avg Unit Price", f"${avg_up:,.2f}")
            # charts
            st.subheader("Monthly trend")
            if df_dash[date_axis_col].notna().any():
                df_dash["_Month"] = df_dash[date_axis_col].dt.to_period("M").dt.to_timestamp()
                ts = df_dash.groupby("_Month", as_index=False).agg({"Quantity":"sum","__Value__":"sum"})
                st.line_chart(ts.set_index("_Month")[["Quantity","__Value__"]])
            else:
                st.info("No valid dates for trend.")
            colA, colB = st.columns(2)
            with colA:
                st.subheader("Top Ship-to Country (Quantity)")
                if "Ship-to Country" in df_dash.columns:
                    top_country = df_dash.groupby("Ship-to Country", as_index=False)["Quantity"].sum().sort_values("Quantity", ascending=False).head(10)
                    st.bar_chart(top_country.set_index("Ship-to Country")["Quantity"])
                else:
                    st.info("Ship-to Country not present.")
            with colB:
                st.subheader("Top Model Name (Quantity)")
                if "Model Name" in df_dash.columns:
                    top_model = df_dash.groupby("Model Name", as_index=False)["Quantity"].sum().sort_values("Quantity", ascending=False).head(10)
                    st.bar_chart(top_model.set_index("Model Name")["Quantity"])
                else:
                    st.info("Model Name not present.")
            st.subheader("Delay code distribution (SAP)")
            cols_delay = [c for c in ["Delay/Early - Confirmation CRD","Delay - PO PSDD Update","Delay - PO PD Update"] if c in df_dash.columns]
            if cols_delay:
                for dc in cols_delay:
                    cnt = df_dash[dc].astype(str).str.upper().replace({"NAN":"","NULL":"","--":"","N/A":""}).str.strip()
                    cnt = cnt[cnt != ""]
                    if len(cnt) == 0:
                        st.info(f"{dc}: no values")
                        continue
                    topd = cnt.value_counts().reset_index()
                    topd.columns = [dc, "Count"]
                    st.bar_chart(topd.set_index(dc)["Count"])
            # downloads
            st.subheader("Download filtered dataset")
            csv_name = f"SalesAnalytics_SAP_Filtered_{today_str_id()}.csv"
            st.download_button("Download CSV", data=df_dash.drop(columns=["__Value__"], errors="ignore").to_csv(index=False).encode("utf-8"), file_name=csv_name)
            if EXCEL_STYLED_AVAILABLE:
                try:
                    bio = io.BytesIO()
                    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
                        df_dash.drop(columns=["__Value__"], errors="ignore").to_excel(writer, index=False, sheet_name="Sales Analytics")
                        ws = writer.sheets["Sales Analytics"]
                        # basic styling
                        header_cells = list(ws.iter_rows(min_row=1, max_row=1, values_only=False))[0]
                        for cell in header_cells:
                            cell.fill = PatternFill("solid", fgColor=OTHER_COLOR)
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                            cell.font = Font(name="Calibri", size=9, bold=True)
                        bio.seek(0)
                    st.download_button("Download Excel (styled)", data=bio.getvalue(), file_name=f"SalesAnalytics_SAP_Filtered_{today_str_id()}.xlsx")
                except Exception as ex:
                    st.warning(f"Styled Excel failed: {ex}")
            # preview
            st.subheader("Preview data (max 1000 rows)")
            st.dataframe(df_dash.drop(columns=["__Value__"], errors="ignore").head(1000), use_container_width=True)
        except Exception as e:
            st.error("Failed to load SAP file for dashboard.")
            st.exception(e)
    else:
        st.info("Upload a SAP Excel to build dashboard.")

# Debug info
with st.expander("🛠 Debug info"):
    try:
        import platform
        st.write("Python:", sys.version)
        st.write("Platform:", platform.platform())
        st.write("Streamlit:", st.__version__)
        st.write("Pandas:", pd.__version__)
        import numpy as np2
        st.write("NumPy:", np2.__version__)
        st.write("openpyxl available:", EXCEL_STYLED_AVAILABLE)
    except Exception as e:
        st.write("Failed to fetch debug info:", e)
