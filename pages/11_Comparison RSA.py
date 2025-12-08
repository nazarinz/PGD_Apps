# app_streamlit.py
# -*- coding: utf-8 -*-
"""
Streamlit wrapper for:
RSA - PGD Comparison Tracking Report pipeline
(keamanan: menjaga semua nama kolom/logic agar sesuai versi asli)
"""

import re
import io
import numpy as np
import pandas as pd
from datetime import datetime
import streamlit as st

# timezone
try:
    from zoneinfo import ZoneInfo
    _tz = ZoneInfo("Asia/Jakarta")
except Exception:
    _tz = None

# =====================
# Helpers (dimuat persis dari script asli)
# =====================
BLANKS = {"(blank)", "blank", "", "--", " -- ", " --"}

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
    """Group-level aggregator untuk Infor: gabung unik tanggal jadi 'MM/DD/YYYY | ...'."""
    dts = to_dt_series(series).dropna()
    if dts.empty:
        return np.nan
    uniq = sorted(set(dts))
    return " | ".join(fmt_dt(x) for x in uniq)

def date_to_text_cell(val) -> str:
    """Row-level SAP display: 1 tanggal → 'MM/DD/YYYY' ('' jika gagal parse)."""
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
    """'11/30/2025 | 12/30/2025' -> {Timestamp('2025-11-30'), ...}; handle single date/iso str."""
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
# Core pipeline as function
# =====================
def run_pipeline(df_sap, df_infor, out_filename=None):
    """
    df_sap: DataFrame loaded from SAP file
    df_infor: DataFrame loaded from Infor file
    returns: df_final (DataFrame), bytes_io (Excel file as bytes)
    """
    # constants & config (mirip skrip asli)
    _today = datetime.now(_tz) if _tz else datetime.now()
    _today_str = _today.strftime("%Y%m%d")
    OUT_JOINED = out_filename or f"RSA - PGD Comparison Tracking Report - {_today_str}.xlsx"

    JOIN_KEYS = ["PO No.(Full)", "CRD_key", "PD_key"]
    DATE_COLS_INFOR = ["Issue Date","FPD","LPD","PSDD","PODD","CRD","PD"]
    DATE_COLS_SAP   = ["Document Date","FPD","LPD","PSDD","PODD","FCR Date","PO Date","CRD","PD","Actual PGI"]

    # =============
    # 2) NORMALISASI INFOR (mirip script)
    # =============
    rename_cols = {
        'Order #': 'PO No.(Full)',
        'Line Aggregator': 'Customer PO item',   # sumber Infor; nanti kita gandakan ke "Line Aggregator"
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
    }
    df_infor = df_infor.rename(columns=rename_cols)

    # standardisasi kolom PD confirmation jika ada
    if "Confirmation Delay Pd" in df_infor.columns and "Delay/Early - Confirmation PD" not in df_infor.columns:
        df_infor = df_infor.rename(columns={"Confirmation Delay Pd": "Delay/Early - Confirmation PD"})

    # buang Quantity==0
    if "Quantity" in df_infor.columns:
        df_infor = df_infor[df_infor["Quantity"].fillna(0) != 0].copy()

    # =====================
    # 3) GROUP INFOR (tanpa 'Customer PO item')
    # =====================
    meta_cols = ["Issue Date", "PO No.(Full)", "Model Name", "Article No", "Ship-to Country", "CRD", "PD"]
    missing = [c for c in meta_cols if c not in df_infor.columns]
    if missing:
        raise ValueError(f"Kolom meta tidak ditemukan di Infor: {missing}")

    # deteksi kolom size (1..18 & variasi)
    size_pat  = re.compile(r'^(?:[1-9]|1[0-8])(?:K|-K|-)?$')
    size_cols = [c for c in df_infor.columns if size_pat.match(str(c))]

    sum_cols   = size_cols + ["Quantity"]
    other_cols = [c for c in df_infor.columns if c not in meta_cols + sum_cols]

    # numeric untuk size & Quantity
    df_inf_num = df_infor.copy()
    if sum_cols:
        df_inf_num[sum_cols] = df_inf_num[sum_cols].apply(pd.to_numeric, errors="coerce")

    # aggregator Infor
    agg_infor = {col: sum_keep_nan for col in sum_cols}
    for col in other_cols:
        if col in DATE_COLS_INFOR:
            agg_infor[col] = date_concat           # multi-date → "A | B"
        else:
            agg_infor[col] = keep_or_join

    df_infor_grouped = (
        df_inf_num.groupby(meta_cols, dropna=False)
                  .agg(agg_infor)
                  .reset_index()
    )

    # buat "Line Aggregator" dari hasil agregasi Customer PO item (jika ada)
    if "Customer PO item" in df_infor_grouped.columns and "Line Aggregator" not in df_infor_grouped.columns:
        df_infor_grouped["Line Aggregator"] = df_infor_grouped["Customer PO item"]

    # kunci merge (datetime-normalized)
    df_infor_grouped["CRD_key"] = to_dt_series(df_infor_grouped["CRD"])
    df_infor_grouped["PD_key"]  = to_dt_series(df_infor_grouped["PD"])

    # =====================
    # 4) SAP (row-level) — display dates as text, keys datetime
    # =====================
    for col in DATE_COLS_SAP:
        if col in df_sap.columns:
            df_sap[col] = df_sap[col].map(date_to_text_cell)

    df_sap["CRD_key"] = to_dt_series(df_sap["CRD"]) if "CRD" in df_sap.columns else pd.NA
    df_sap["PD_key"]  = to_dt_series(df_sap["PD"]) if "PD" in df_sap.columns else pd.NA

    # =====================
    # 5) MERGE (JOIN_KEYS) + Prefix kolom Infor
    # =====================
    infor_cols_for_merge = [
        "Order Status","Article No","LPD","PODD","PSDD","FPD","CRD","PD",
        "Delay/Early - Confirmation CRD","Delay - PO PSDD Update","Delay - PO PD Update",
        "Quantity","Shipment Method","Issue Date",
        "Customer PO item","Line Aggregator"
    ]
    if "Delay/Early - Confirmation PD" in df_infor_grouped.columns:
        infor_cols_for_merge = ["Delay/Early - Confirmation PD"] + infor_cols_for_merge

    inf_pick_cols = [c for c in infor_cols_for_merge if c in df_infor_grouped.columns]
    inf_pick = df_infor_grouped[JOIN_KEYS + inf_pick_cols].copy()
    pref_map = {c: f"infor {c}" for c in inf_pick_cols}
    inf_pick = inf_pick.rename(columns=pref_map)

    df_join = df_sap.merge(inf_pick, on=JOIN_KEYS, how="left")

    # =====================
    # 6) Mapping delay code Infor → standar
    # =====================
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

    # =====================
    # 7) COMPARE (num/str/delay) + tanggal sebagai set
    # =====================
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

    # Quanity (SAP) vs infor Quantity -> Result_Quantity
    if "Quanity" in df.columns and "infor Quantity" in df.columns:
        df["Result_Quantity"] = equal_series(norm_num(df["Quanity"]), norm_num(df["infor Quantity"])).fillna(False)

    # Article No
    if "Article No" in df.columns and "infor Article No" in df.columns:
        df["Result Article No"] = equal_series(norm_str(df["Article No"]), norm_str(df["infor Article No"])).fillna(False)

    # Delay fields (after mapping)
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

    # Pastikan "Line Aggregator" ada (fallback dari infor)
    if "Line Aggregator" not in df.columns and "infor Line Aggregator" in df.columns:
        df["Line Aggregator"] = df["infor Line Aggregator"]

    # =====================
    # 8) REORDER persis sesuai template user
    # =====================
    desired_order = [
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
    present = [c for c in desired_order if c in df.columns]
    rest     = [c for c in df.columns if c not in present]
    df_final = df[present + rest]

    # =====================
    # 9) EXPORT EXCEL — shortdate per-cell utk kolom tanggal
    # =====================
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

    # write to BytesIO
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
            # rewrite kolom tanggal per-cell
            for cidx, col in enumerate(df_final.columns):
                if col in date_display_cols:
                    for ridx, val in enumerate(df_final[col].tolist(), start=1):  # +1 karena header di row 0
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

            # conditional format untuk semua kolom Result*
            res_cols = [c for c in df_final.columns if c.startswith("Result ")]
            for col in res_cols:
                cidx = df_final.columns.get_loc(col)
                rng  = f"{to_excel_col(cidx)}2:{to_excel_col(cidx)}{nrows+1}"
                ws.conditional_format(rng, {"type":"cell","criteria":"==","value":"TRUE","format":fmt_bool_T})
                ws.conditional_format(rng, {"type":"cell","criteria":"==","value":"FALSE","format":fmt_bool_F})

            # set lebar kolom
            for idx, col in enumerate(df_final.columns, start=1):
                if col.startswith("Result "):
                    ws.set_column(idx-1, idx-1, 16)
                elif col in date_display_cols:
                    ws.set_column(idx-1, idx-1, 12)  # tanggal
                else:
                    ws.set_column(idx-1, idx-1, 18)

            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, nrows, ncols-1)
            writer.save()
        out.seek(0)
    except Exception as e:
        # fallback openpyxl: simpler write (tanggal tidak per-cell)
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            df_final.to_excel(writer, index=False, sheet_name="Data")
        out.seek(0)

    return df_final, out, OUT_JOINED

# =====================
# Streamlit UI
# =====================
st.set_page_config(page_title="RSA - PGD Comparison Pipeline", layout="wide")

st.title("RSA - PGD Comparison Tracking — Streamlit")
st.caption("Versi Streamlit dari pipeline (preserve semua logic & kolom).")

with st.sidebar:
    st.markdown("## Input Files")
    st.markdown("Upload file SAP dan Infor. Format: .xlsx, .xls, .csv.")
    sap_file = st.file_uploader("Upload SAP file (ZRSD...)", type=["xlsx","xls","csv"], key="sap")
    infor_file = st.file_uploader("Upload Infor file (Book1...)", type=["xlsx","xls","csv"], key="infor")
    st.markdown("---")
    st.markdown("Jika kosong, Anda dapat memasukkan path lokal di server (untuk advanced).")
    local_sap_path = st.text_input("Path lokal SAP (optional)", value="", help="Contoh: /content/ZRSD1013 - 20251204 RSA.XLSX")
    local_infor_path = st.text_input("Path lokal Infor (optional)", value="", help="Contoh: /content/Book1.xlsx")
    st.markdown("---")
    run_btn = st.button("Run pipeline", type="primary")

st.info("Pipeline akan mempertahankan semua nama kolom asli — jangan ubah header file kecuali tahu konsekuensinya.")

# preview helpers
def load_excel_or_csv(uploaded):
    if uploaded is None:
        return None
    try:
        if hasattr(uploaded, "read"):
            # streamlit upload: BytesIO
            uploaded.seek(0)
            # try excel first
            try:
                df = pd.read_excel(uploaded)
                return df
            except Exception:
                uploaded.seek(0)
                df = pd.read_csv(uploaded)
                return df
        else:
            # path string
            path = str(uploaded)
            if path.lower().endswith((".xls",".xlsx")):
                return pd.read_excel(path)
            else:
                return pd.read_csv(path)
    except Exception as e:
        st.error(f"Load error: {e}")
        return None

# load inputs (priority: uploaded files > local path)
df_sap = None
df_infor = None

if sap_file is not None:
    df_sap = load_excel_or_csv(sap_file)

if infor_file is not None:
    df_infor = load_excel_or_csv(infor_file)

# local fallback if not uploaded
if df_sap is None and local_sap_path:
    try:
        df_sap = load_excel_or_csv(local_sap_path)
    except Exception as e:
        st.warning(f"Gagal load local SAP: {e}")

if df_infor is None and local_infor_path:
    try:
        df_infor = load_excel_or_csv(local_infor_path)
    except Exception as e:
        st.warning(f"Gagal load local Infor: {e}")

# show previews
col1, col2 = st.columns(2)
with col1:
    st.subheader("Preview SAP")
    if df_sap is not None:
        st.dataframe(df_sap.head(200))
    else:
        st.info("Belum ada SAP file ter-load.")

with col2:
    st.subheader("Preview Infor")
    if df_infor is not None:
        st.dataframe(df_infor.head(200))
    else:
        st.info("Belum ada Infor file ter-load.")

# run
if run_btn:
    if df_sap is None or df_infor is None:
        st.error("File SAP dan Infor harus terisi (upload atau path lokal).")
    else:
        with st.spinner("Running pipeline — please wait..."):
            try:
                df_result, bytes_io, out_name = run_pipeline(df_sap.copy(), df_infor.copy())
                st.success("Pipeline selesai.")
                st.download_button(
                    label=f"Download Excel: {out_name}",
                    data=bytes_io.getvalue(),
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.markdown("### Hasil (preview)")
                st.dataframe(df_result.head(200))
                # show some simple stats
                st.markdown("### Ringkasan kolom Result*")
                res_cols = [c for c in df_result.columns if c.startswith("Result ")]
                if res_cols:
                    summary = {c: int(df_result[c].sum(skipna=True)) for c in res_cols if df_result[c].dtype == bool or df_result[c].dtype == 'object'}
                    st.json(summary)
                else:
                    st.info("Tidak ada kolom Result* di hasil (cek mapping kolom input).")
            except Exception as e:
                st.exception(e)
                st.error("Pipeline gagal — periksa file input dan header kolom yang diperlukan.")

st.markdown("---")
st.caption("Aplikasi ini membaca header persis seperti versi skrip. Bila ada error 'Kolom meta tidak ditemukan', periksa header Infor: 'Issue Date','PO No.(Full)','Model Name','Article No','Ship-to Country','CRD','PD'.")
