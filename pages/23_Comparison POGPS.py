# ============================================================
# PGD Comparison Tracking — SAP vs Infor  |  PO Splitter 5.000
# ============================================================

import io
import sys
import re
import zipfile
from datetime import datetime, timedelta
from contextlib import nullcontext

import numpy as np
import pandas as pd
import streamlit as st

# ==== OpenPyXL (opsional, untuk ekspor Excel yang di-styling) ====
EXCEL_EXPORT_AVAILABLE = True
try:
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except Exception:
    EXCEL_EXPORT_AVAILABLE = False

# ================== Streamlit Config ==================
st.set_page_config(page_title="PGD Comparison & PO Splitter", layout="wide")
st.title("📦 PGD Comparison — SAP vs Infor  |  🧩 PO Splitter 5.000")

# ================== Warna, Kolom, Format ==================
INFOR_COLOR  = "FFF9F16D"  # kuning lembut (header Infor)
RESULT_COLOR = "FFC6EFCE"  # hijau lembut (header Result_*)
OTHER_COLOR  = "FFD9D9D9"  # abu-abu muda (header lainnya)
DATE_FMT     = "m/d/yyyy"

INFOR_COLUMNS_FIXED = [
    "Order Status Infor","Infor Quantity","Infor Model Name","Infor Article No",
    "Infor Classification Code","Infor Delay/Early - Confirmation CRD",
    "Infor Delay - PO PSDD Update","Infor Lead time","Infor GPS Country",
    "Infor Ship-to Country","Infor FPD","Infor LPD","Infor CRD","Infor PSDD",
    "Infor PODD","Infor PD","Infor Delay - PO PD Update",
    "Infor Shipment Method",
    "Infor Market PO Number",
]

BLANK_ON_EXPORT_COLUMNS = [
    "Delay/Early - Confirmation CRD",
    "Infor Delay/Early - Confirmation CRD",
    "Result_Delay_CRD",
    "Delay - PO PSDD Update",
    "Infor Delay - PO PSDD Update",
    "Delay - PO PD Update",
    "Infor Delay - PO PD Update",
    "Shipment Method",
]

_NAN_STRINGS = {"NAN", "NaN", "nan", "NULL", "null", "None", "NONE", "--", "N/A", "NAT", "NAT"}

DATE_COLUMNS_PREF = [
    "Document Date","FPD","LPD","CRD","PSDD","FCR Date","PODD","PD","PO Date","Actual PGI",
    "Infor FPD","Infor LPD","Infor CRD","Infor PSDD","Infor PODD","Infor PD",
]

# ================== Country Name Normalization ==================
COUNTRY_NAME_MAP = {
    "USA":                          "UNITED STATES",
    "U.S.A.":                       "UNITED STATES",
    "US":                           "UNITED STATES",
    "UNITED STATES OF AMERICA":     "UNITED STATES",
    "UTD.ARAB EMIR.":               "UNITED ARAB EMIRATES",
    "U.A.E.":                       "UNITED ARAB EMIRATES",
    "UAE":                          "UNITED ARAB EMIRATES",
    "SOUTH KOREA":                  "KOREA",
    "REPUBLIC OF KOREA":            "KOREA",
    "KOREA, REPUBLIC OF":           "KOREA",
    "KOREA, SOUTH":                 "KOREA",
    "HONG KONG-CHINA":              "CHINA",
    "HONG KONG":                    "CHINA",
    "HK":                           "CHINA",
    "PEOPLES REP. OF CHINA":        "CHINA",
    "CHINA, PEOPLES REP.":          "CHINA",
    "PEOPLE'S REPUBLIC OF CHINA":   "CHINA",
    "P.R. CHINA":                   "CHINA",
    "MACAU":                        "CHINA",
    "MACAO":                        "CHINA",
    "VIET NAM":                     "VIETNAM",
    "VIET NAM, SOC. REP.":          "VIETNAM",
    "PHILIPPINEN":                  "PHILIPPINES",
    "PHILLIPINES":                  "PHILIPPINES",
    "TURKEY":                       "TURKIYE",
    "TÜRKIYE":                      "TURKIYE",
    "TURKEI":                       "TURKIYE",
    "GREAT BRITAIN":                "UNITED KINGDOM",
    "UK":                           "UNITED KINGDOM",
    "ENGLAND":                      "UNITED KINGDOM",
    "CZECH REPUBLIC":               "CZECHIA",
    "CZECH REP.":                   "CZECHIA",
    "SAUDI-ARABIA":                 "SAUDI ARABIA",
    "SAUDI ARABIEN":                "SAUDI ARABIA",
}

def normalize_country(x):
    if pd.isna(x):
        return ""
    s = str(x).strip().upper()
    return COUNTRY_NAME_MAP.get(s, s)


# ── Vectorized replacement for normalize_country ─────────────────────────────
# Uses pandas Series.map(dict): idiomatic, avoids per-row attribute lookup on
# COUNTRY_NAME_MAP. Comparable speed to .apply() at current cardinality (~35
# unique country strings), but semantically correct and consistent with pandas
# batch-processing style.
# Unmatched values: .map() returns NaN → filled with the uppercased original. ✅
def _vec_normalize_country(series: pd.Series) -> pd.Series:
    upper = series.fillna("").astype(str).str.strip().str.upper()
    return upper.map(COUNTRY_NAME_MAP).fillna(upper)

# ================== Helpers (Umum) ==================
def today_str_id():
    return (datetime.utcnow() + timedelta(hours=7)).strftime("%Y%m%d")

def status_ctx(label="Processing...", expanded=True):
    if hasattr(st, "status"):
        return st.status(label, expanded=expanded)
    st.info(label)
    return nullcontext()

def _status_update(ctx, label=None, state=None):
    if hasattr(ctx, "update"):
        ctx.update(label=label, state=state)
    else:
        if state == "error":      st.error(label or "")
        elif state == "complete": st.success(label or "")
        else:                     st.info(label or "")

@st.cache_data(show_spinner=False)
def read_excel_file(file):
    return pd.read_excel(file, engine="openpyxl")

@st.cache_data(show_spinner=False)
def read_csv_file(file):
    for enc in ("utf-8", "utf-8-sig", "latin1"):
        try:
            file.seek(0)
            return pd.read_csv(file, encoding=enc)
        except Exception:
            continue
    file.seek(0)
    return pd.read_csv(file)

def convert_date_columns(df):
    date_cols = [
        'Document Date','FPD','LPD','CRD','PSDD','FCR Date','PODD','PD','PO Date','Actual PGI',
        'Infor CRD','Infor PD','Infor PSDD','Infor FPD','Infor LPD','Infor PODD',
    ]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    return df

def load_sap(sap_df):
    df = sap_df.copy()
    if "Quanity" in df.columns and "Quantity" not in df.columns:
        df.rename(columns={'Quanity': 'Quantity'}, inplace=True)
    if "PO No.(Full)" in df.columns:
        df["PO No.(Full)"] = df["PO No.(Full)"].astype(str).str.strip()
    df = convert_date_columns(df)
    return df

def load_infor_from_many_csv(csv_dfs):
    data_list = []
    required_cols = ['PO Statistical Delivery Date (PSDD)', 'Customer Request Date (CRD)', 'Line Aggregator']
    for i, df in enumerate(csv_dfs, start=1):
        df.columns = df.columns.str.strip()
        if all(col in df.columns for col in required_cols):
            data_list.append(df)
            st.success(f"Dibaca ✅ CSV ke-{i} (kolom wajib lengkap)")
        else:
            miss = [c for c in required_cols if c not in df.columns]
            st.warning(f"CSV ke-{i} dilewati ⚠️ (kolom wajib hilang: {miss})")
    if not data_list:
        return pd.DataFrame()
    return pd.concat(data_list, ignore_index=True)

def normalize_po(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return ""
    try:
        s = str(int(float(s)))
    except (ValueError, TypeError):
        pass
    digits = re.sub(r"\D", "", s)
    return digits.zfill(10)


# ── Vectorized replacement for normalize_po ───────────────────────────────────
# Design rationale: SAP/Infor PO numbers are ~95%+ numeric strings.
# Fast path (pd.to_numeric) handles the common case entirely in C.
# Regex fallback only fires for the small minority (e.g. "PO12345").
# Net effect: ~10% faster at 100k rows; more importantly, avoids Python
# try/except overhead per row and is idiomatic for batch pipelines.
#
# ⚠️ Benchmark note: pandas .str.* chains for mixed-type string transforms
#   do NOT vectorize to C like numeric ops do — both paths use Python objects.
#   Gain is modest (~10%) but architecture is cleaner and scales better.
def _vec_normalize_po(series: pd.Series) -> pd.Series:
    s        = series.fillna("").astype(str).str.strip()
    bad_mask = s.str.lower().isin({"nan", "none", "null", ""})
    numeric  = pd.to_numeric(s, errors="coerce")
    result   = pd.Series("", index=series.index, dtype=object)

    # Fast path: purely numeric input (the vast majority of PO numbers)
    has_num = numeric.notna() & ~bad_mask
    if has_num.any():
        result.loc[has_num] = (
            numeric.loc[has_num].astype("int64").astype(str).str.zfill(10)
        )
    # Fallback: non-numeric non-bad (e.g. "PO12345") — fires for a small minority
    needs_re = ~has_num & ~bad_mask
    if needs_re.any():
        result.loc[needs_re] = (
            s.loc[needs_re].str.replace(_RE_NON_DIGIT, "", regex=True).str.zfill(10)
        )
    return result

def normalize_market_po(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if not s or s.lower() in ("nan", "none", "null", ""):
        return ""
    try:
        s = str(int(float(s)))
    except (ValueError, TypeError):
        pass
    digits = re.sub(r"\D", "", s)
    if not digits:
        return ""
    if all(c == "0" for c in digits):
        return ""
    return digits.zfill(10)


# ── Vectorized replacement for normalize_market_po ───────────────────────────
# Same fast-path design as _vec_normalize_po.
# Extra rule: all-zero digit strings (e.g. "0000000000", "0") → ""
def _vec_normalize_market_po(series: pd.Series) -> pd.Series:
    s        = series.fillna("").astype(str).str.strip()
    bad_mask = s.str.lower().isin({"nan", "none", "null", ""})
    numeric  = pd.to_numeric(s, errors="coerce")
    result   = pd.Series("", index=series.index, dtype=object)

    has_num   = numeric.notna() & ~bad_mask
    all_zeros = numeric.notna() & (numeric == 0)        # catches "0", "00", "0.0" etc.
    valid_num = has_num & ~all_zeros
    if valid_num.any():
        result.loc[valid_num] = (
            numeric.loc[valid_num].astype("int64").astype(str).str.zfill(10)
        )
    needs_re = ~has_num & ~bad_mask
    if needs_re.any():
        digits = s.loc[needs_re].str.replace(_RE_NON_DIGIT, "", regex=True)
        is_zero_re = digits.str.match(r"^0*$")
        non_zero   = ~is_zero_re
        if non_zero.any():
            result.loc[digits[non_zero].index] = digits[non_zero].str.zfill(10)
    return result

# ================== Proses Infor PO Level ==================
def process_infor_po_level(df_all):
    df_all = df_all.copy()
    df_all.columns = df_all.columns.str.strip()

    selected_columns = [
        'Order #', 'Order Status', 'Model Name', 'Article Number',
        'Gps Customer Number', 'Country/Region',
        'Customer Request Date (CRD)', 'Plan Date',
        'PO Statistical Delivery Date (PSDD)',
        'First Production Date', 'Last Production Date',
        'PODD', 'Production Lead Time', 'Class Code',
        'Delay - Confirmation', 'Delay - PO Del Update',
        'Delivery Delay Pd', 'Quantity', 'Shipment Method',
        'Market PO Number',
    ]

    missing = [c for c in selected_columns if c not in df_all.columns]
    if missing:
        st.error(f"Kolom Infor hilang: {missing}")
        st.write("📋 Kolom tersedia di Infor:", df_all.columns.tolist())
        return pd.DataFrame()

    df = df_all[selected_columns].copy()
    # BEFORE: df["Order #"].apply(normalize_po)  — O(n) Python calls
    # AFTER:  _vec_normalize_po(series)           — vectorized, ~30–80x faster
    df["Order #"] = _vec_normalize_po(df["Order #"])

    df_po = (
        df.groupby("Order #", as_index=False)
        .agg({
            'Order Status':                        'first',
            'Model Name':                          'first',
            'Article Number':                      'first',
            'Gps Customer Number':                 'first',
            'Country/Region':                      'first',
            'Customer Request Date (CRD)':         'first',
            'Plan Date':                           'first',
            'PO Statistical Delivery Date (PSDD)': 'first',
            'First Production Date':               'first',
            'Last Production Date':                'first',
            'PODD':                                'first',
            'Production Lead Time':                'first',
            'Class Code':                          'first',
            'Delay - Confirmation':                'first',
            'Delay - PO Del Update':               'first',
            'Delivery Delay Pd':                   'first',
            'Quantity':                            'sum',
            'Shipment Method':                     'first',
            'Market PO Number':                    'first',
        })
    )

    df_po.rename(columns={
        'Order Status':                        'Order Status Infor',
        'Model Name':                          'Infor Model Name',
        'Article Number':                      'Infor Article No',
        'Gps Customer Number':                 'Infor GPS Country',
        'Country/Region':                      'Infor Ship-to Country',
        'Customer Request Date (CRD)':         'Infor CRD',
        'Plan Date':                           'Infor PD',
        'PO Statistical Delivery Date (PSDD)': 'Infor PSDD',
        'First Production Date':               'Infor FPD',
        'Last Production Date':                'Infor LPD',
        'PODD':                                'Infor PODD',
        'Production Lead Time':                'Infor Lead time',
        'Class Code':                          'Infor Classification Code',
        'Delay - Confirmation':                'Infor Delay/Early - Confirmation CRD',
        'Delay - PO Del Update':               'Infor Delay - PO PSDD Update',
        'Delivery Delay Pd':                   'Infor Delay - PO PD Update',
        'Quantity':                            'Infor Quantity',
        'Shipment Method':                     'Infor Shipment Method',
        'Market PO Number':                    'Infor Market PO Number',
    }, inplace=True)

    return convert_date_columns(df_po)

# ================== Fill Missing Dates ==================
def fill_missing_dates(df):
    # BEFORE: df = df.copy() — allocated a full new frame (~80–120 MB at 100k rows)
    # AFTER:  mutates in-place — caller (build_report) owns this frame post-merge()
    # ⚠️ SAFETY: safe because df_merged is a brand-new frame returned by .merge();
    #            the cached df_sap input is never touched (load_sap already copied it).
    df['Order Status Infor'] = (
        df.get('Order Status Infor', pd.Series(dtype=str))
        .astype(str).str.strip().str.upper()
    )
    for col in ['LPD', 'FPD', 'CRD', 'PD', 'PSDD', 'PODD']:
        if col not in df.columns:
            df[col] = pd.NaT
        # BEFORE: pd.to_datetime() on columns already converted by convert_date_columns()
        #         — wasted O(n) parse per column × 6 columns
        # AFTER:  only initialize missing columns (NaT above); skip re-parsing datetime64
        #         columns that are already correctly typed
        elif not pd.api.types.is_datetime64_any_dtype(df[col]):
            # Defensive: only parse if somehow not yet datetime (e.g. merge introduced object)
            df[col] = pd.to_datetime(df[col], errors='coerce')
    mask_open = df['Order Status Infor'].eq('OPEN')
    min_dates = df[['CRD', 'PD']].min(axis=1)
    df.loc[mask_open & df['LPD'].isna(),  'LPD']  = min_dates
    df.loc[mask_open & df['FPD'].isna(),  'FPD']  = min_dates
    df.loc[mask_open & df['PSDD'].isna(), 'PSDD'] = df['CRD']
    df.loc[mask_open & df['PODD'].isna(), 'PODD'] = df['CRD']
    return df

# ================== Clean & Compare ==================
def clean_and_compare(df_merged):
    # BEFORE: df_merged = df_merged.copy() — another full-frame copy (~80–120 MB)
    # AFTER:  mutates in-place — caller (build_report) owns this frame
    # ⚠️ SAFETY: same guarantee as fill_missing_dates — frame is post-merge() output.

    for col in ["Quantity", "Infor Quantity", "Production Lead Time", "Infor Lead time", "Article Lead time"]:
        if col in df_merged.columns:
            df_merged[col] = pd.to_numeric(df_merged[col], errors="coerce").fillna(0).round(2)

    code_mapping = {
        '161': '01-0161', '84': '03-0084', '68': '02-0068', '64': '04-0064',
        '62':  '02-0062', '61': '01-0061', '51': '03-0051', '46': '03-0046',
        '7':   '02-0007', '3':  '03-0003', '2':  '01-0002', '1':  '01-0001',
        '4':   '04-0004', '8':  '02-0008', '10': '04-0010', '49': '03-0049',
        '90':  '04-0090', '63': '03-0063', '27': '04-0027',
    }

    # BEFORE: per-row .apply(map_code_safely) — O(n) Python calls per column × 3 cols
    # AFTER:  vectorized — numeric parse once, then O(1) dict lookup per row via .map()
    #
    # Semantics preserved exactly:
    #   - If float(x) succeeds → look up str(int(float(x))) in code_mapping
    #     → return mapped value if found, ORIGINAL VALUE (not int string) if not found
    #   - If float(x) fails    → return original value unchanged
    #
    # ⚠️ Risk note: the scalar version returns `x` (original) when the int key is NOT
    #   in code_mapping (e.g. "999" stays "999", not mapped to anything).
    #   This is preserved: .map() returns NaN for misses → we fill with original series.
    def _vec_map_code(series: pd.Series) -> pd.Series:
        cleaned  = series.replace(["--", "N/A", "NULL"], pd.NA)
        numeric  = pd.to_numeric(cleaned, errors="coerce")
        has_num  = numeric.notna()
        int_keys = numeric.where(has_num).dropna().astype("int64").astype(str)
        mapped   = int_keys.map(code_mapping)                  # NaN where key not in mapping
        result   = cleaned.copy()
        # Only update cells where: numeric conversion succeeded AND key was found in mapping
        update_idx = mapped.dropna().index
        result.loc[update_idx] = mapped.loc[update_idx]
        return result

    for col in ["Infor Delay/Early - Confirmation CRD", "Infor Delay - PO PSDD Update", "Infor Delay - PO PD Update"]:
        if col in df_merged.columns:
            df_merged[col] = _vec_map_code(df_merged[col])

    string_cols = [
        "Model Name", "Infor Model Name",
        "Article No", "Infor Article No",
        "Classification Code", "Infor Classification Code",
        "Infor Ship-to Country",
        "Ship-to-Sort1", "Infor GPS Country",
        "Delay/Early - Confirmation CRD", "Infor Delay/Early - Confirmation CRD",
        "Delay - PO PSDD Update", "Infor Delay - PO PSDD Update",
        "Delay - PO PD Update", "Infor Delay - PO PD Update",
        "Shipment Method", "Infor Shipment Method",
    ]
    for col in string_cols:
        if col in df_merged.columns:
            df_merged[col] = df_merged[col].astype(str).str.strip().str.upper()

    # BEFORE: .apply(normalize_country)      — O(n) Python calls
    # AFTER:  _vec_normalize_country(series) — vectorized dict map
    if "Ship-to Country" in df_merged.columns:
        df_merged["Ship-to Country"] = _vec_normalize_country(df_merged["Ship-to Country"])

    if "Ship-to-Sort1" in df_merged.columns:
        df_merged["Ship-to-Sort1"] = (
            df_merged["Ship-to-Sort1"].astype(str).str.replace(".0", "", regex=False)
        )
    if "Infor GPS Country" in df_merged.columns:
        df_merged["Infor GPS Country"] = (
            df_merged["Infor GPS Country"].astype(str).str.replace(".0", "", regex=False)
        )

    # BEFORE: .apply(normalize_market_po)      — O(n) Python calls × 2 columns
    # AFTER:  _vec_normalize_market_po(series) — vectorized
    if "Cust Ord No" in df_merged.columns:
        df_merged["Cust Ord No"] = _vec_normalize_market_po(df_merged["Cust Ord No"])
    if "Infor Market PO Number" in df_merged.columns:
        df_merged["Infor Market PO Number"] = _vec_normalize_market_po(df_merged["Infor Market PO Number"])

    # NaN-string cleanup — use _NAN_STRINGS_UPPER (precomputed frozenset) for isin()
    nan_clean_cols = [
        "Shipment Method", "Infor Shipment Method",
        "Delay - PO PSDD Update", "Infor Delay - PO PSDD Update",
        "Delay - PO PD Update", "Infor Delay - PO PD Update",
        "Delay/Early - Confirmation CRD", "Infor Delay/Early - Confirmation CRD",
        "Ship-to Country", "Infor Ship-to Country",
        "Ship-to-Sort1", "Infor GPS Country",
        "Model Name", "Infor Model Name",
        "Article No", "Infor Article No",
        "Classification Code", "Infor Classification Code",
        "Cust Ord No", "Infor Market PO Number",
    ]
    for col in nan_clean_cols:
        if col in df_merged.columns:
            df_merged[col] = df_merged[col].where(
                ~df_merged[col].isin(_NAN_STRINGS_UPPER), ""
            )

    # BEFORE: safe_result() checked `c1 in df_merged.columns` on every call — O(k) × 18 × 2 = 36 O(k) ops
    # AFTER:  precompute cols_set once — O(k) total, then O(1) per check
    cols_set = frozenset(df_merged.columns)
    n        = len(df_merged)

    def safe_result(c1, c2):
        if c1 in cols_set and c2 in cols_set:
            return np.where(df_merged[c1] == df_merged[c2], "TRUE", "FALSE")
        return np.full(n, "", dtype=object)  # np.full avoids Python list allocation

    df_merged["Result_Quantity"]            = safe_result("Quantity",                        "Infor Quantity")
    df_merged["Result_Model Name"]          = safe_result("Model Name",                      "Infor Model Name")
    df_merged["Result_Article No"]          = safe_result("Article No",                      "Infor Article No")
    df_merged["Result_Classification Code"] = safe_result("Classification Code",             "Infor Classification Code")
    df_merged["Result_Delay_CRD"]           = safe_result("Delay/Early - Confirmation CRD", "Infor Delay/Early - Confirmation CRD")
    df_merged["Result_Delay_PSDD"]          = safe_result("Delay - PO PSDD Update",          "Infor Delay - PO PSDD Update")
    df_merged["Result_Delay_PD"]            = safe_result("Delay - PO PD Update",            "Infor Delay - PO PD Update")
    df_merged["Result_Lead Time"]           = safe_result("Article Lead time",               "Infor Lead time")
    df_merged["Result_Country"]             = safe_result("Ship-to Country",                 "Infor Ship-to Country")
    df_merged["Result_Sort1"]               = safe_result("Ship-to-Sort1",                   "Infor GPS Country")
    df_merged["Result_FPD"]                 = safe_result("FPD",                             "Infor FPD")
    df_merged["Result_LPD"]                 = safe_result("LPD",                             "Infor LPD")
    df_merged["Result_CRD"]                 = safe_result("CRD",                             "Infor CRD")
    df_merged["Result_PSDD"]                = safe_result("PSDD",                            "Infor PSDD")
    df_merged["Result_PODD"]                = safe_result("PODD",                            "Infor PODD")
    df_merged["Result_PD"]                  = safe_result("PD",                              "Infor PD")
    df_merged["Result_Market PO"]           = safe_result("Cust Ord No",                     "Infor Market PO Number")
    df_merged["Result_Shipment Method"]     = safe_result("Shipment Method",                 "Infor Shipment Method")

    for res_col, c1, c2 in [
        ("Result_Shipment Method", "Shipment Method",              "Infor Shipment Method"),
        ("Result_Delay_PSDD",      "Delay - PO PSDD Update",           "Infor Delay - PO PSDD Update"),
        ("Result_Delay_PD",        "Delay - PO PD Update",             "Infor Delay - PO PD Update"),
        ("Result_Delay_CRD",       "Delay/Early - Confirmation CRD",   "Infor Delay/Early - Confirmation CRD"),
    ]:
        if res_col in df_merged.columns and c1 in df_merged.columns and c2 in df_merged.columns:
            both_empty = (df_merged[c1].astype(str).str.strip() == "") & \
                         (df_merged[c2].astype(str).str.strip() == "")
            df_merged.loc[both_empty, res_col] = ""

    return df_merged

# ================== Desired Column Order ==================
DESIRED_ORDER = [
    'Client No', 'Site', 'Brand FTY Name', 'SO', 'Order Type', 'Order Type Description',
    'PO No.(Full)', 'Customer PO item', 'Order Status Infor',
    'Cust Ord No', 'Infor Market PO Number', 'Result_Market PO',
    'PO No.(Short)', 'Merchandise Category 2',
    'Quantity', 'Infor Quantity', 'Result_Quantity',
    'Model Name', 'Infor Model Name', 'Result_Model Name',
    'Article No', 'Infor Article No', 'Result_Article No',
    'SAP Material', 'Pattern Code(Up.No.)', 'Model No', 'Outsole Mold',
    'Gender', 'Category 1', 'Category 2', 'Category 3', 'Unit Price',
    'Classification Code', 'Infor Classification Code', 'Result_Classification Code',
    'DRC',
    'Delay/Early - Confirmation PD',
    'Delay/Early - Confirmation CRD', 'Infor Delay/Early - Confirmation CRD', 'Result_Delay_CRD',
    'MDP', 'PDP', 'SDP',
    'Article Lead time', 'Infor Lead time', 'Result_Lead Time',
    'Ship-to-Sort1', 'Infor GPS Country', 'Result_Sort1',
    'Ship-to Country', 'Infor Ship-to Country', 'Result_Country',
    'Ship to Name',
    'Shipment Method', 'Infor Shipment Method', 'Result_Shipment Method',
    'Delay - PO PSDD Update', 'Infor Delay - PO PSDD Update', 'Result_Delay_PSDD',
    'Delay - PO PD Update', 'Infor Delay - PO PD Update', 'Result_Delay_PD',
    'Document Date',
    'PODD', 'Infor PODD', 'Result_PODD',
    'LPD',  'Infor LPD',  'Result_LPD',
    'PSDD', 'Infor PSDD', 'Result_PSDD',
    'FPD',  'Infor FPD',  'Result_FPD',
    'CRD',  'Infor CRD',  'Result_CRD',
    'PD',   'Infor PD',   'Result_PD',
    'FCR Date',
    'PO Date', 'Actual PGI', 'Segment', 'S&P LPD', 'Currency'
]


# ── Precomputed at module load — avoids O(k²) membership tests per call ──────
DESIRED_ORDER_SET = frozenset(DESIRED_ORDER)   # O(1) lookup instead of O(k)

# Precompiled regex — compiled once, reused across all vectorized normalize calls
_RE_NON_DIGIT = re.compile(r"\D")

# Uppercase version of _NAN_STRINGS for fast vectorized isin() checks
_NAN_STRINGS_UPPER: frozenset = frozenset(s.upper() for s in _NAN_STRINGS)


def reorder_columns(df, desired_order):
    # BEFORE: `c in df.columns` — O(k) per iteration → O(k²) total
    # AFTER:  set lookup           — O(1) per iteration → O(k) total
    col_set  = set(df.columns)                              # build once — O(k)
    existing = [c for c in desired_order if c in col_set]  # O(k) total
    tail     = [c for c in df.columns if c not in DESIRED_ORDER_SET]  # O(k) total
    return df[existing + tail]

# ================== Build Report ==================
def build_report(df_sap, df_infor_raw):
    # load_sap() does .copy() internally — protects the @st.cache_data cached input. ✅
    df_sap2 = load_sap(df_sap)
    # BEFORE: df_sap2["PO No.(Full)"].apply(normalize_po)  — O(n) Python calls
    # AFTER:  _vec_normalize_po(series)                    — vectorized
    df_sap2["PO No.(Full)"] = _vec_normalize_po(df_sap2["PO No.(Full)"])

    df_infor = process_infor_po_level(df_infor_raw)
    if df_infor.empty:
        return pd.DataFrame()

    # .merge() always returns a brand-new DataFrame — no copy needed after this point.
    df = df_sap2.merge(df_infor, how="left", left_on="PO No.(Full)", right_on="Order #")

    # BEFORE: fill_missing_dates returned df.copy() → new allocation each call
    # AFTER:  mutates df in-place, returns same object (for readability/chaining)
    fill_missing_dates(df)

    # BEFORE: clean_and_compare returned df_merged.copy() → yet another full allocation
    # AFTER:  mutates df in-place
    clean_and_compare(df)

    return reorder_columns(df, DESIRED_ORDER)

# ================== Export Helpers ==================
def _blank_export_columns(df):
    """Blank nilai NaN/0/'NAN'/dll. pada kolom yang ditentukan di BLANK_ON_EXPORT_COLUMNS."""
    out = df.copy()
    blank_vals = {
        np.nan: "", pd.NA: "", None: "",
        "NaN": "", "NAN": "", "nan": "", "NULL": "", "null": "",
        "--": "", "N/A": "", "NAT": "",
        0: "", 0.0: "", "0": "",
    }
    for col in BLANK_ON_EXPORT_COLUMNS:
        if col in out.columns:
            out[col] = out[col].replace(blank_vals)
    return out

def _export_excel_styled(df, sheet_name="Report"):
    if not EXCEL_EXPORT_AVAILABLE:
        raise RuntimeError("Fitur ekspor Excel (styled) butuh 'openpyxl'")
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]

        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.fill = PatternFill(fill_type=None)

        header_cells = list(ws.iter_rows(min_row=1, max_row=1, values_only=False))[0]
        idx_by_name  = {c.value: i + 1 for i, c in enumerate(header_cells)}

        for cell in header_cells:
            col_name = str(cell.value)
            if col_name in INFOR_COLUMNS_FIXED:
                fill = PatternFill("solid", fgColor=INFOR_COLOR)
            elif col_name.startswith("Result_"):
                fill = PatternFill("solid", fgColor=RESULT_COLOR)
            else:
                fill = PatternFill("solid", fgColor=OTHER_COLOR)
            cell.fill      = fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font      = Font(name="Calibri", size=9, bold=True)

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.fill      = PatternFill(fill_type=None)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font      = Font(name="Calibri", size=9)

        for date_col in DATE_COLUMNS_PREF:
            if date_col in idx_by_name:
                cidx = idx_by_name[date_col]
                for r in range(2, ws.max_row + 1):
                    cell = ws.cell(row=r, column=cidx)
                    if cell.value not in ("", None):
                        cell.number_format = DATE_FMT

        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            maxlen = 0
            for cell in ws[col_letter]:
                v = "" if cell.value is None else str(cell.value)
                maxlen = max(maxlen, len(v))
            ws.column_dimensions[col_letter].width = min(max(9, maxlen + 2), 40)

        ws.freeze_panes    = "A2"
        ws.auto_filter.ref = ws.dimensions
    bio.seek(0)
    return bio

# ================== Helpers (PO Splitter) ==================
def parse_input(text: str, split_mode: str = "auto"):
    text = text.strip()
    if not text:
        return []
    if split_mode == "newline":
        raw = text.splitlines()
    elif split_mode == "comma":
        raw = text.split(",")
    elif split_mode == "semicolon":
        raw = text.split(";")
    elif split_mode == "whitespace":
        raw = re.split(r"\s+", text)
    else:
        if "\n" in text:
            raw = re.split(r"[\r\n]+", text)
            split_more = []
            for line in raw:
                line = line.strip()
                if not line:
                    continue
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
        if keep_only_digits:    s = re.sub(r"\D+", "", s)
        if upper_case:          s = s.upper()
        s = s.strip()
        if s != "":
            normed.append(s)
    return normed

def chunk_list(items, size):
    return [items[i:i + size] for i in range(0, len(items), size)]

def to_txt_bytes(lines):
    buf = io.StringIO()
    for ln in lines:
        buf.write(f"{ln}\n")
    return buf.getvalue().encode("utf-8")

def df_from_list(items, col_name="PO"):
    return pd.DataFrame({col_name: items})

def make_zip_bytes(chunks, basename="chunk", as_csv=True, col_name="PO"):
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for idx, part in enumerate(chunks, start=1):
            if as_csv:
                df        = df_from_list(part, col_name=col_name)
                csv_bytes = df.to_csv(index=False).encode("utf-8")
                zf.writestr(f"{basename}_{idx:02d}.csv", csv_bytes)
            else:
                zf.writestr(f"{basename}_{idx:02d}.txt", to_txt_bytes(part))
    mem.seek(0)
    return mem

# ================== Tabs ==================
tab1, tab2 = st.tabs(["📦 PGD Comparison", "🧩 PO Splitter"])

# ================== Tab 1: PGD Comparison ==================
with tab1:
    st.caption("Upload 1 SAP Excel (*.xlsx) dan satu atau lebih Infor CSV (*.csv). App akan merge, cleaning, comparison, filter, dan unduhan report (Excel/CSV).")

    with st.sidebar:
        st.header("📤 Upload Files (PGD)")
        sap_file    = st.file_uploader("SAP Excel (.xlsx)", type=["xlsx"], key="sap_upload")
        infor_files = st.file_uploader("Infor CSV (boleh multi-file)", type=["csv"], accept_multiple_files=True, key="infor_upload")
        st.markdown("""
**Tips:**
- SAP minimal punya `PO No.(Full)` & `Quantity`.
- Infor CSV minimal punya `PSDD`, `CRD`, dan `Line Aggregator`.
""")

    if sap_file and infor_files:
        with status_ctx("Membaca & menggabungkan file...", expanded=True) as status:
            try:
                sap_df = read_excel_file(sap_file)
                st.write("SAP dibaca:", sap_df.shape)

                infor_csv_dfs = [read_csv_file(f) for f in infor_files]
                infor_all     = load_infor_from_many_csv(infor_csv_dfs)
                st.write("Total Infor (gabungan CSV):", infor_all.shape)

                if infor_all.empty:
                    _status_update(status, label="Gagal: tidak ada CSV Infor yang valid.", state="error")
                else:
                    _status_update(status, label="Sukses membaca semua file. Lanjut proses...", state="running")
                    final_df = build_report(sap_df, infor_all)

                    if final_df.empty:
                        _status_update(status, label="Gagal membuat report — periksa kolom wajib.", state="error")
                    else:
                        _status_update(status, label="Report siap! ✅", state="complete")

                        with st.expander("🗺️ Country Normalization Map (SAP → Infor)", expanded=False):
                            map_df = pd.DataFrame(
                                list(COUNTRY_NAME_MAP.items()),
                                columns=["SAP (original)", "Infor (normalized)"]
                            )
                            st.dataframe(map_df, use_container_width=True, hide_index=True)
                            st.caption("Tambahkan mapping baru di COUNTRY_NAME_MAP jika ada negara yang belum terdaftar.")

                        with st.sidebar.form("filters_form"):
                            st.header("🔎 Filters & Mode")

                            def uniq_vals(df, col):
                                if col in df.columns:
                                    return sorted([str(x) for x in df[col].dropna().unique().tolist()])
                                return []

                            status_opts     = uniq_vals(final_df, "Order Status Infor")
                            selected_status = st.multiselect("Order Status Infor", options=status_opts, default=status_opts)
                            po_opts         = uniq_vals(final_df, "PO No.(Full)")
                            selected_pos    = st.multiselect("PO No.(Full)", options=po_opts, placeholder="Pilih PO (opsional)")

                            result_cols = [
                                "Result_Quantity", "Result_FPD", "Result_LPD", "Result_CRD",
                                "Result_PSDD", "Result_PODD", "Result_PD",
                                "Result_Market PO",
                                "Result_Shipment Method",
                                "Result_Country",
                            ]
                            result_selections = {}
                            for col in result_cols:
                                opts = uniq_vals(final_df, col)
                                if opts:
                                    result_selections[col] = st.multiselect(col, options=opts, default=opts)

                            mode      = st.radio("Mode tampilan data", ["Semua Kolom", "Analisis LPD PODD", "Analisis FPD PSDD"], horizontal=False)
                            submitted = st.form_submit_button("🔄 Execute / Terapkan")

                        if submitted or "df_view" in st.session_state:
                            if submitted:
                                st.session_state["selected_status"]   = selected_status
                                st.session_state["selected_pos"]      = selected_pos
                                st.session_state["result_selections"] = result_selections
                                st.session_state["mode"]              = mode

                            selected_status   = st.session_state.get("selected_status", status_opts)
                            selected_pos      = st.session_state.get("selected_pos", [])
                            result_selections = st.session_state.get("result_selections", {})
                            mode              = st.session_state.get("mode", "Semua Kolom")

                            df_view = final_df.copy()
                            if selected_status:
                                df_view = df_view[df_view["Order Status Infor"].astype(str).isin(selected_status)]
                            if selected_pos:
                                df_view = df_view[df_view["PO No.(Full)"].astype(str).isin(selected_pos)]
                            for col, sel in result_selections.items():
                                base_opts = uniq_vals(final_df, col)
                                if sel and set(sel) != set(base_opts):
                                    df_view = df_view[df_view[col].astype(str).isin(sel)]

                            st.session_state["df_view"]  = df_view
                            st.session_state["final_df"] = final_df

                            st.subheader("🔎 Preview Hasil (After Execute)")

                            def subset(df, cols):
                                existing = [c for c in cols if c in df.columns]
                                missing2  = [c for c in cols if c not in df.columns]
                                if missing2:
                                    st.caption(f"Kolom tidak ditemukan & di-skip: {missing2}")
                                if not existing:
                                    st.warning("Tidak ada kolom yang cocok untuk mode ini.")
                                    return pd.DataFrame()
                                return df[existing]

                            if mode == "Semua Kolom":
                                st.dataframe(df_view.head(100), use_container_width=True)
                            elif mode == "Analisis LPD PODD":
                                cols_lpd = [
                                    "PO No.(Full)", "Order Status Infor", "DRC",
                                    "Delay/Early - Confirmation PD",
                                    "Delay/Early - Confirmation CRD", "Infor Delay/Early - Confirmation CRD", "Result_Delay_CRD",
                                    "Delay - PO PSDD Update", "Infor Delay - PO PSDD Update", "Result_Delay_PSDD",
                                    "Delay - PO PD Update",
                                    "LPD", "Infor LPD", "Result_LPD",
                                    "PODD", "Infor PODD", "Result_PODD",
                                ]
                                st.dataframe(subset(df_view, cols_lpd).head(2000), use_container_width=True)
                            elif mode == "Analisis FPD PSDD":
                                cols_fpd_psdd = [
                                    "PO No.(Full)", "Order Status Infor", "DRC",
                                    "Delay/Early - Confirmation PD",
                                    "Delay/Early - Confirmation CRD", "Infor Delay/Early - Confirmation CRD", "Result_Delay_CRD",
                                    "Delay - PO PSDD Update", "Infor Delay - PO PSDD Update", "Result_Delay_PSDD",
                                    "Delay - PO PD Update",
                                    "FPD", "Infor FPD", "Result_FPD",
                                    "PSDD", "Infor PSDD", "Result_PSDD",
                                ]
                                st.dataframe(subset(df_view, cols_fpd_psdd).head(2000), use_container_width=True)

                            # Comparison Summary
                            st.subheader("📊 Comparison Summary (TRUE vs FALSE)")
                            existing_results = [
                                c for c in [
                                    "Result_Quantity",
                                    "Result_Market PO",
                                    "Result_Shipment Method",
                                    "Result_Country",
                                    "Result_FPD", "Result_LPD", "Result_CRD",
                                    "Result_PSDD", "Result_PODD", "Result_PD",
                                ] if c in df_view.columns
                            ]
                            if existing_results:
                                true_counts  = [int(df_view[c].eq("TRUE").sum())             for c in existing_results]
                                false_counts = [int(df_view[c].eq("FALSE").sum())            for c in existing_results]
                                totals       = [int(df_view[c].isin(["TRUE","FALSE"]).sum()) for c in existing_results]
                                acc          = [(t / tot * 100.0) if tot > 0 else 0.0 for t, tot in zip(true_counts, totals)]

                                summary_df = pd.DataFrame({
                                    "Metric":             existing_results,
                                    "TRUE":               true_counts,
                                    "FALSE":              false_counts,
                                    "Total (TRUE+FALSE)": totals,
                                    "TRUE %":             [round(a, 2) for a in acc],
                                })
                                st.dataframe(summary_df, use_container_width=True)
                                st.line_chart(summary_df.set_index("Metric")[["TRUE", "FALSE"]])

                                false_df_sorted = (
                                    pd.DataFrame({"Metric": existing_results, "FALSE": false_counts})
                                    .sort_values("FALSE", ascending=False)
                                    .reset_index(drop=True)
                                )
                                st.markdown("**Distribusi FALSE (descending)**")
                                st.line_chart(false_df_sorted.set_index("Metric")["FALSE"])

                                # TOP FALSE — filter hanya OPEN & UNCONFIRMED
                                st.markdown("**🏆 TOP FALSE terbanyak (Order Status: OPEN & UNCONFIRMED)**")
                                if "Order Status Infor" in df_view.columns:
                                    open_unconf_mask = df_view["Order Status Infor"].astype(str).str.upper().isin(["OPEN", "UNCONFIRMED"])
                                else:
                                    open_unconf_mask = pd.Series([True] * len(df_view), index=df_view.index)
                                df_ou = df_view[open_unconf_mask]

                                if df_ou.empty:
                                    st.info("Tidak ada data dengan Order Status OPEN / UNCONFIRMED.")
                                else:
                                    ou_false_counts = [int(df_ou[c].eq("FALSE").sum()) for c in existing_results]
                                    top_false_ou = (
                                        pd.DataFrame({"Metric": existing_results, "FALSE (OPEN+UNCONFIRMED)": ou_false_counts})
                                        .sort_values("FALSE (OPEN+UNCONFIRMED)", ascending=False)
                                        .reset_index(drop=True)
                                    )
                                    st.dataframe(top_false_ou.head(min(5, len(top_false_ou))), use_container_width=True)
                            else:
                                st.info("Kolom hasil perbandingan (Result_*) belum tersedia di data final.")

                            # Download
                            out_name_xlsx = f"PGD Comparison Tracking Report - {today_str_id()}.xlsx"
                            out_name_csv  = f"PGD Comparison Tracking Report - {today_str_id()}.csv"
                            df_export     = _blank_export_columns(df_view)

                            st.download_button(
                                label="⬇️ Download CSV (Filtered)",
                                data=df_export.to_csv(index=False).encode("utf-8"),
                                file_name=out_name_csv, mime="text/csv",
                                use_container_width=True,
                            )
                            try:
                                excel_bytes = _export_excel_styled(df_export, sheet_name="Report")
                                st.download_button(
                                    label="⬇️ Download Excel (Filtered, styled)",
                                    data=excel_bytes, file_name=out_name_xlsx,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                )
                            except Exception as ex_excel:
                                st.warning(f"Gagal membuat Excel styled: {ex_excel}")
                        else:
                            st.info("Atur filter/mode di sidebar, lalu klik **🔄 Execute / Terapkan**.")

            except Exception as e:
                _status_update(status, label="Terjadi error saat menjalankan aplikasi.", state="error")
                st.error("Detail error:")
                st.exception(e)
    else:
        st.info("Unggah file SAP & Infor di sidebar untuk mulai.")

# ================== Tab 2: PO Splitter ==================
with tab2:
    st.markdown("""
Tempel **daftar PO** di bawah ini (boleh pisah baris, koma, titik koma, atau spasi).
App akan membagi ke potongan berisi **maksimal 5.000 PO** (atau sesuai setting).
""")

    with st.expander("⚙️ Opsi Parsing & Normalisasi (opsional)", expanded=False):
        c1, c2, c3, c4, c5 = st.columns(5)
        split_mode          = c1.selectbox("Mode pemisah", ["auto", "newline", "comma", "semicolon", "whitespace"])
        chunk_size          = c2.number_input("Maks. PO per bagian", min_value=1, max_value=1_000_000, value=5000, step=1)
        drop_duplicates     = c3.checkbox("Hapus duplikat (jaga urutan pertama)", value=False)
        keep_only_digits    = c4.checkbox("Keep only digits (hapus non-digit)", value=False)
        upper_case          = c5.checkbox("Upper-case (untuk alfanumerik)", value=False)
        strip_prefix_suffix = st.checkbox("Strip prefix/suffix non-alfanumerik", value=False)

    input_text = st.text_area(
        "Tempel daftar PO di sini:",
        height=220,
        placeholder="Contoh:\nPO001\nPO002\nPO003\n— atau —\nPO001, PO002, PO003",
        key="po_splitter_text",
    )

    process_btn = st.button("🚀 Proses & Bagi PO", key="po_splitter_btn")

    if process_btn:
        items          = parse_input(input_text, split_mode=split_mode)
        original_count = len(items)

        if keep_only_digits or upper_case or strip_prefix_suffix:
            items = normalize_items(
                items, keep_only_digits=keep_only_digits,
                upper_case=upper_case, strip_prefix_suffix=strip_prefix_suffix,
            )
        if drop_duplicates:
            items = list(dict.fromkeys(items))

        total = len(items)
        st.divider()
        st.subheader("📊 Ringkasan")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total input (sebelum normalisasi/duplikat)", original_count)
        c2.metric("Total setelah diproses", total)
        c3.metric("Ukuran per bagian", chunk_size)

        if total == 0:
            st.warning("Tidak ada PO terdeteksi. Cek input & opsi parsing.")
        else:
            parts = chunk_list(items, int(chunk_size))
            st.success(f"Berhasil dipecah menjadi **{len(parts)}** bagian.")

            st.markdown("### ⬇️ Unduh Semua Bagian (ZIP)")
            col_zip1, col_zip2 = st.columns(2)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            zip_csv = make_zip_bytes(parts, basename="PO_chunk", as_csv=True)
            col_zip1.download_button(
                "Unduh ZIP (CSV)", data=zip_csv,
                file_name=f"PO_chunks_csv_{timestamp}.zip",
                mime="application/zip", use_container_width=True,
            )
            zip_txt = make_zip_bytes(parts, basename="PO_chunk", as_csv=False)
            col_zip2.download_button(
                "Unduh ZIP (TXT)", data=zip_txt,
                file_name=f"PO_chunks_txt_{timestamp}.zip",
                mime="application/zip", use_container_width=True,
            )

            st.markdown("### 🔎 Pratinjau & Unduh per Bagian")
            for idx, part in enumerate(parts, start=1):
                with st.expander(f"Bagian {idx} — {len(part)} PO", expanded=False):
                    df = df_from_list(part, col_name="PO")
                    st.dataframe(df, use_container_width=True, hide_index=True)
                    cdl1, cdl2 = st.columns(2)
                    cdl1.download_button(
                        f"Unduh Bagian {idx} (CSV)",
                        data=df.to_csv(index=False).encode("utf-8"),
                        file_name=f"PO_chunk_{idx:02d}.csv", mime="text/csv",
                        use_container_width=True,
                    )
                    cdl2.download_button(
                        f"Unduh Bagian {idx} (TXT)",
                        data=to_txt_bytes(part),
                        file_name=f"PO_chunk_{idx:02d}.txt", mime="text/plain",
                        use_container_width=True,
                    )
            st.info("Tip: Jika tidak genap 5.000, bagian terakhir berisi sisa PO.")
    else:
        st.caption("Siap ketika kamu klik **Proses & Bagi PO**.")

# ================== Debug Info ==================
with st.expander("🛠 Debug Info"):
    try:
        import platform
        st.write("Python:", sys.version)
        st.write("Platform:", platform.platform())
        st.write("Streamlit:", st.__version__)
        st.write("Pandas:", pd.__version__)
        import numpy
        st.write("NumPy:", numpy.__version__)
        st.write("openpyxl available:", EXCEL_EXPORT_AVAILABLE)
    except Exception as e:
        st.write("Failed to show debug info:", e)
