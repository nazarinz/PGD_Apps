# app.py
# Streamlit front-end for PGD — Merge per-PO (Exploded) + Dashboard vs SAP Compare (No Aggregation)
# Versi: 2025-10-21 (Streamlit)

import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime, timedelta
import io

st.set_page_config(page_title="PGD Merge & Compare (Exploded)", layout="wide")

st.title("PGD — Merge per-PO (Exploded) + Dashboard vs SAP Compare")
st.caption("Pairing 1:1 nearest · CRD Monthly · MDP/SDP Gap (NZ) · Customer PO item & Line Aggregator dari SAP")

# Custom CSS dengan Dark Mode Support
st.markdown("""
<style>
:root {
    --primary: #1f77b4;
    --primary-light: #5b9ee1;
    --text-primary: #333;
    --text-secondary: #666;
    --bg-card: #ffffff;
    --bg-secondary: #f8f9fa;
    --border: #e0e0e0;
}

@media (prefers-color-scheme: dark) {
    :root {
        --primary: #5b9ee1;
        --primary-light: #7fb3f0;
        --text-primary: #e0e0e0;
        --text-secondary: #a0a0a0;
        --bg-card: #2d2d2d;
        --bg-secondary: #1e1e1e;
        --border: #404040;
    }
}

.analytics-container {
    background: var(--bg-secondary);
    padding: 1.5rem;
    border-radius: 10px;
    margin: 1.5rem 0;
    border: 1px solid var(--border);
}

.metric-box {
    background: var(--bg-card);
    padding: 1.5rem;
    border-radius: 8px;
    border: 1px solid var(--border);
    margin: 0.5rem 0;
    transition: all 0.3s ease;
    box-shadow: 0 2px 6px rgba(0, 0, 0, 0.08);
}

.metric-box:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(31, 119, 180, 0.15);
}

.metric-label {
    color: var(--text-secondary);
    font-size: 0.85rem;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-bottom: 0.5rem;
}

.metric-value {
    color: var(--primary-light);
    font-size: 1.8rem;
    font-weight: bold;
}

.analytics-header {
    background: linear-gradient(135deg, var(--primary) 0%, #0d47a1 100%);
    color: white;
    padding: 2rem;
    border-radius: 10px;
    margin-bottom: 2rem;
    box-shadow: 0 4px 12px rgba(31, 119, 180, 0.2);
}

.analytics-header h1 {
    margin: 0;
    color: white;
    text-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
}

.chart-container {
    background: var(--bg-card);
    padding: 1.5rem;
    border-radius: 10px;
    border: 1px solid var(--border);
    margin: 1.5rem 0;
    box-shadow: 0 2px 6px rgba(0, 0, 0, 0.08);
}

.stats-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 1rem;
    margin: 1.5rem 0;
}

.stat-card {
    background: var(--bg-card);
    padding: 1.5rem;
    border-radius: 8px;
    border: 1px solid var(--border);
    text-align: center;
    transition: all 0.3s ease;
}

.stat-card:hover {
    border-color: var(--primary);
    box-shadow: 0 4px 12px rgba(31, 119, 180, 0.15);
}

* {
    transition: background-color 0.3s ease, color 0.3s ease, border-color 0.3s ease;
}

@media (max-width: 768px) {
    .analytics-header {
        padding: 1rem;
    }
    
    .analytics-header h1 {
        font-size: 1.5rem;
    }
    
    .stats-grid {
        grid-template-columns: 1fr;
    }
}
</style>
""", unsafe_allow_html=True)

# -------------------------
# Helpers (same logic as script)
# -------------------------
def _clean_cols(pdf: pd.DataFrame) -> pd.DataFrame:
    pdf = pdf.copy()
    pdf.columns = (
        pdf.columns
        .astype(str)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    return pdf

def _ref_qty_sap_row(row: pd.Series):
    iq = pd.to_numeric(row.get("Infor Quantity"), errors="coerce")
    q  = pd.to_numeric(row.get("Quantity"), errors="coerce")
    return iq if pd.notna(iq) else q

def compare_numeric(df, left, right):
    a, b = pd.to_numeric(df.get(left), errors="coerce"), pd.to_numeric(df.get(right), errors="coerce")
    return (a.notna() & b.notna() & (a == b))

def compare_date(df, left, right):
    a, b = pd.to_datetime(df.get(left), errors="coerce"), pd.to_datetime(df.get(right), errors="coerce")
    return (a.notna() & b.notna() & (a.dt.date == b.dt.date))

def _norm_str_col(s):
    return s.astype("string").str.strip().str.upper()

def build_pairs_for_po(df_sap_po: pd.DataFrame, df_dash_po: pd.DataFrame):
    sap_idx = list(df_sap_po.index)
    dash_idx = list(df_dash_po.index)

    if len(sap_idx) == 0 and len(dash_idx) == 0:
        return []
    if len(sap_idx) == 0:
        return [(None, j) for j in dash_idx]
    if len(dash_idx) == 0:
        return [(i, None) for i in sap_idx]

    sap_ref = pd.Series({_i: _ref_qty_sap_row(df_sap_po.loc[_i]) for _i in sap_idx})
    dash_qty = pd.to_numeric(df_dash_po["Dashboard Quantity"], errors="coerce")

    candidates = []
    pos_map = {i: p for p, i in enumerate(sap_idx)}
    for i in sap_idx:
        ref = sap_ref[i]
        for j in dash_idx:
            dq = dash_qty.loc[j]
            if pd.isna(ref) or pd.isna(dq):
                dist = np.inf
            else:
                dist = abs(float(ref) - float(dq))
            candidates.append((dist, - (dq if pd.notna(dq) else -np.inf), pos_map[i], i, j))

    candidates.sort(key=lambda x: (x[0], x[1], x[2]))
    used_sap, used_dash, chosen = set(), set(), []
    for dist, _negdq, _ord, i, j in candidates:
        if i in used_sap or j in used_dash:
            continue
        chosen.append((i, j))
        used_sap.add(i); used_dash.add(j)

    for i in sap_idx:
        if i not in used_sap:
            chosen.append((i, None))
    for j in dash_idx:
        if j not in used_dash:
            chosen.append((None, j))

    return chosen

# -------------------------
# UI - uploads
# -------------------------
with st.sidebar:
    st.markdown("## Upload files")
    pbi_file = st.file_uploader("Upload PBI Data (dashboard) — Excel", type=["xlsx","xls"], key="pbi")
    sap_file = st.file_uploader("Upload SAP/INFOR report — Excel", type=["xlsx","xls"], key="sap")
    st.markdown("---")
    st.markdown("Options:")
    show_preview = st.checkbox("Tampilkan preview df_final setelah proses", value=True)
    run_button = st.button("Jalankan pipeline & buat file")

# -------------------------
# Processing
# -------------------------
def process(pbi_file, sap_file):
    PK = "PO Number (GPS)"

    # load
    df_db = pd.read_excel(pbi_file, engine="openpyxl")
    df_db = _clean_cols(df_db)
    df_sapinf = pd.read_excel(sap_file, engine="openpyxl")
    df_sapinf = _clean_cols(df_sapinf)

    # normalize PBI columns (drop exact duplicates & alias .n)
    df = df_db.copy()
    df = df.loc[:, ~df.columns.duplicated()].copy()
    alias_cols = [c for c in df.columns if re.search(r"\.\d+$", str(c))]
    for c in alias_cols:
        base = re.sub(r"\.\d+$", "", c)
        if base in df.columns:
            df[base] = df[base].combine_first(df[c])
            df.drop(columns=[c], inplace=True)
        else:
            df.rename(columns={c: base}, inplace=True)

    # light rename
    dashboard_alias = {
        "Elevated check": "Elevated Check",
        "Elevated_Check": "Elevated Check",
        "Responsiveness Status": "Responsiveness",
        "Responsiveness_": "Responsiveness",
        "PO Status (Dashboard)": "PO Status",
    }
    df.rename(columns={k:v for k,v in dashboard_alias.items() if k in df.columns}, inplace=True)

    # ensure join key exists
    if PK not in df.columns:
        raise KeyError(f"Kolom PK '{PK}' tidak ada di PBI Data.")
    sapinf_key_candidates = ["PO Number (GPS)", "PO No.(Full)", "PO No.", "PO Number"]
    join_left_key = next((k for k in sapinf_key_candidates if k in df_sapinf.columns), None)
    if join_left_key is None:
        raise KeyError(f"Tidak menemukan key join di df_sapinf. Coba salah satu: {sapinf_key_candidates}")

    df_sapinf_join = df_sapinf.copy()
    if join_left_key != PK:
        df_sapinf_join[PK] = df_sapinf_join[join_left_key]

    # prepare dashboard columns (no aggregation)
    dash_map_src_to_new = {
        "Order ALL":              "Dashboard Quantity",
        "MDP Status Adjusted":    "Dashboard MDP Status Adjusted",
        "SDP Status Adjusted":    "Dashboard SDP Status Adjusted",
        "FPD_":                   "Dashboard FPD",
        "LPD_":                   "Dashboard LPD",
        "CRD_":                   "Dashboard CRD",
        "PSDD_":                  "Dashboard PSDD",
        "Planned Date_":          "Dashboard PD",
        "PODD_":                  "Dashboard PODD",
        "FGR Document Date_":     "Dashboard FGR",
    }
    present_src = [src for src in dash_map_src_to_new if src in df.columns]
    passthrough_cols = ["Elevated Check", "Responsiveness", "PO Status"]
    keep_cols = [PK] + present_src + passthrough_cols
    df_dash_keep = df[keep_cols].copy(deep=True)
    df_dash_keep = df_dash_keep.rename(columns={src: dash_map_src_to_new[src] for src in present_src})
    for src, newc in dash_map_src_to_new.items():
        if newc not in df_dash_keep.columns:
            df_dash_keep[newc] = pd.NA
    for c in passthrough_cols:
        if c not in df_dash_keep.columns:
            df_dash_keep[c] = pd.NA

    # pairing per PO
    out_rows = []
    for po, df_sap_po in df_sapinf_join.groupby(PK, dropna=False):
        df_dash_po = df_dash_keep[df_dash_keep[PK] == po]
        pairs = build_pairs_for_po(df_sap_po, df_dash_po)
        sap_ref_row = df_sap_po.iloc[0] if len(df_sap_po) > 0 else None

        sap_single = (len(df_sap_po) == 1) and (len(df_dash_po) > 1)
        best_dash_idx_for_single = None
        if sap_single and sap_ref_row is not None:
            rq = _ref_qty_sap_row(sap_ref_row)
            if pd.notna(rq) and "Dashboard Quantity" in df_dash_po.columns:
                d_diffs = (pd.to_numeric(df_dash_po["Dashboard Quantity"], errors="coerce") - rq).abs()
                if not d_diffs.dropna().empty:
                    best_dash_idx_for_single = d_diffs.idxmin()

        for i, j in pairs:
            if i is not None:
                row_sap = df_sap_po.loc[i]
                out = row_sap.to_dict()
            else:
                out = (sap_ref_row.copy().to_dict() if sap_ref_row is not None else {})
                out["Quantity"] = np.nan
                out["Infor Quantity"] = np.nan

            if j is not None:
                row_dash = df_dash_po.loc[j]
            else:
                row_dash = pd.Series(dtype="object")

            for col in [
                "Dashboard Quantity","Dashboard MDP Status Adjusted","Dashboard SDP Status Adjusted",
                "Dashboard FPD","Dashboard LPD","Dashboard CRD","Dashboard PSDD",
                "Dashboard PD","Dashboard PODD","Dashboard FGR",
                "Elevated Check","Responsiveness","PO Status"
            ]:
                out[col] = row_dash.get(col, pd.NA)

            # Ambil Customer PO item & Line Aggregator dari SAP (raw)
            for sap_col in ["Customer PO item", "Line Aggregator"]:
                if i is not None:
                    out[sap_col] = row_sap.get(sap_col, pd.NA)
                else:
                    out[sap_col] = (sap_ref_row.get(sap_col, pd.NA) if sap_ref_row is not None else pd.NA)

            if sap_single and (j is not None) and (j != best_dash_idx_for_single):
                out["Quantity"] = np.nan
                out["Infor Quantity"] = np.nan

            out_rows.append(out)

    df_out = pd.DataFrame(out_rows)

    # keep raw types for Customer PO item & Line Aggregator (do not cast)
    for c in ["Elevated Check", "Responsiveness", "PO Status"]:
        if c in df_out.columns:
            df_out[c] = df_out[c].astype("string").str.strip()

    for c in ["Quantity", "Infor Quantity", "Dashboard Quantity"]:
        if c in df_out.columns:
            df_out[c] = pd.to_numeric(df_out[c], errors="coerce")

    # compare columns
    if ("Quantity" in df_out.columns) and ("Infor Quantity" in df_out.columns):
        df_out["Result_Quantity"] = compare_numeric(df_out, "Quantity", "Infor Quantity")
    else:
        df_out["Result_Quantity"] = pd.NA

    if ("Quantity" in df_out.columns) and ("Dashboard Quantity" in df_out.columns):
        df_out["Dashboard vs SAP Result_Quantity"] = compare_numeric(df_out, "Quantity", "Dashboard Quantity")
    else:
        df_out["Dashboard vs SAP Result_Quantity"] = pd.NA

    date_compare_pairs = [
        ("FPD",  "Infor FPD",  "Dashboard FPD",  "Result_FPD",  "Dashboard vs SAP Result_FPD"),
        ("LPD",  "Infor LPD",  "Dashboard LPD",  "Result_LPD",  "Dashboard vs SAP Result_LPD"),
        ("CRD",  "Infor CRD",  "Dashboard CRD",  "Result_CRD",  "Dashboard vs SAP Result_CRD"),
        ("PSDD", "Infor PSDD", "Dashboard PSDD", "Result_PSDD", "Dashboard vs SAP Result_PSDD"),
        ("PODD", "Infor PODD", "Dashboard PODD", "Result_PODD", "Dashboard vs SAP Result_PODD"),
        ("PD",   "Infor PD",   "Dashboard PD",   "Result_PD",   "Dashboard vs SAP Result_PD"),
    ]

    for sap_col, infor_col, dash_col, result_col, dashvs_col in date_compare_pairs:
        if (sap_col in df_out.columns) and (infor_col in df_out.columns):
            df_out[result_col] = compare_date(df_out, sap_col, infor_col)
        else:
            df_out[result_col] = pd.NA
        if (sap_col in df_out.columns) and (dash_col in df_out.columns):
            df_out[dashvs_col] = compare_date(df_out, sap_col, dash_col)
        else:
            df_out[dashvs_col] = pd.NA

    if ("FCR Date" in df_out.columns) and ("Dashboard FGR" in df_out.columns):
        df_out["Dashboard vs SAP Result FCR"] = compare_date(df_out, "FCR Date", "Dashboard FGR")
    else:
        df_out["Dashboard vs SAP Result FCR"] = pd.NA

    # MDP/SDP gap logic
    for need in ["MDP","SDP","Dashboard MDP Status Adjusted","Dashboard SDP Status Adjusted",
                 "Quantity","Dashboard Quantity"]:
        if need not in df_out.columns:
            df_out[need] = pd.NA

    qty_sap_num   = pd.to_numeric(df_out["Quantity"], errors="coerce").fillna(0)
    qty_dash_num  = pd.to_numeric(df_out["Dashboard Quantity"], errors="coerce").fillna(0)

    mdp_status    = _norm_str_col(df_out["MDP"])
    dash_mdp_stat = _norm_str_col(df_out["Dashboard MDP Status Adjusted"])
    sdp_status    = _norm_str_col(df_out["SDP"])
    dash_sdp_stat = _norm_str_col(df_out["Dashboard SDP Status Adjusted"])

    mask_mdp_fail         = mdp_status.eq("FAIL").fillna(False).to_numpy()
    mask_dash_mdp_delay   = dash_mdp_stat.eq("DELAY").fillna(False).to_numpy()
    mask_sdp_fail         = sdp_status.eq("FAIL").fillna(False).to_numpy()
    mask_dash_sdp_delay   = dash_sdp_stat.eq("DELAY").fillna(False).to_numpy()

    df_out["MDP Delay Qty"]           = np.where(mask_mdp_fail,       qty_sap_num,  0)
    df_out["Dashboard MDP Delay Qty"] = np.where(mask_dash_mdp_delay, qty_dash_num, 0)
    df_out["GAP MDP"]                 = df_out["Dashboard MDP Delay Qty"] - df_out["MDP Delay Qty"]

    df_out["SDP Delay Qty"]           = np.where(mask_sdp_fail,       qty_sap_num,  0)
    df_out["Dashboard SDP Delay Qty"] = np.where(mask_dash_sdp_delay, qty_dash_num, 0)
    df_out["GAP SDP"]                 = df_out["Dashboard SDP Delay Qty"] - df_out["SDP Delay Qty"]

    for c in ["MDP Delay Qty","Dashboard MDP Delay Qty","GAP MDP",
              "SDP Delay Qty","Dashboard SDP Delay Qty","GAP SDP"]:
        df_out[c] = pd.to_numeric(df_out[c], errors="coerce").round(0).astype("Int64")

    gap_mdp_num = pd.to_numeric(df_out["GAP MDP"], errors="coerce")
    gap_sdp_num = pd.to_numeric(df_out["GAP SDP"], errors="coerce")
    df_out["GAP MDP (NZ)"] = gap_mdp_num.where(gap_mdp_num.ne(0)).round(0).astype("Int64")
    df_out["GAP SDP (NZ)"] = gap_sdp_num.where(gap_sdp_num.ne(0)).round(0).astype("Int64")
    df_out["Has MDP Gap"] = gap_mdp_num.fillna(0).ne(0)
    df_out["Has SDP Gap"] = gap_sdp_num.fillna(0).ne(0)

    # dates & crd monthly
    date_cols_should = [
        "Document Date", "FPD", "Infor FPD", "Dashboard FPD",
        "LPD", "Infor LPD", "Dashboard LPD",
        "CRD", "Infor CRD", "Dashboard CRD",
        "PSDD", "Infor PSDD", "Dashboard PSDD",
        "FCR Date", "Dashboard FGR",
        "PODD", "Infor PODD", "Dashboard PODD",
        "PD", "Infor PD", "Dashboard PD",
        "PO Date", "Actual PGI"
    ]
    for c in [x for x in date_cols_should if x in df_out.columns]:
        df_out[c] = pd.to_datetime(df_out[c], errors="coerce")

    if "CRD" in df_out.columns:
        crd_dt = pd.to_datetime(df_out["CRD"], errors="coerce")
        crd_mon = crd_dt.dt.strftime("%Y%m").where(crd_dt.notna(), pd.NA)
        df_out["CRD Monthly"] = crd_mon
    else:
        df_out["CRD Monthly"] = pd.NA

    # final order (trimmed to required columns; omitted ones you asked to remove)
    final_order = [
        "Client No","Site","Brand FTY Name","SO","Order Type","Order Type Description",
        "PO No.(Full)","Customer PO item","Line Aggregator",
        "Elevated Check","Responsiveness","PO Status","Order Status Infor","PO No.(Short)","Merchandise Category 2",
        "Quantity","Infor Quantity","Dashboard Quantity","Result_Quantity","Dashboard vs SAP Result_Quantity",
        "Model Name",
        "Article No","Infor Article No","Result_Article No",
        "SAP Material","Pattern Code(Up.No.)","Model No","Outsole Mold","Gender",
        "Category 1","Category 2","Category 3","Unit Price",
        "DRC","Delay/Early - Confirmation PD","Delay/Early - Confirmation CRD",
        "Infor Delay/Early - Confirmation CRD","Result_Delay_CRD",
        "Delay - PO PSDD Update","Infor Delay - PO PSDD Update","Result_Delay_PSDD",
        "Delay - PO PD Update","Infor Delay - PO PD Update","Result_Delay_PD",
        "MDP","Dashboard MDP Status Adjusted",
        "MDP Delay Qty","Dashboard MDP Delay Qty","GAP MDP","GAP MDP (NZ)","Has MDP Gap",
        "PDP",
        "SDP","Dashboard SDP Status Adjusted",
        "SDP Delay Qty","Dashboard SDP Delay Qty","GAP SDP","GAP SDP (NZ)","Has SDP Gap",
        "Document Date","FPD","Infor FPD","Dashboard FPD","Result_FPD","Dashboard vs SAP Result_FPD",
        "LPD","Infor LPD","Dashboard LPD","Result_LPD","Dashboard vs SAP Result_LPD",
        "CRD Monthly","CRD","Infor CRD","Dashboard CRD","Result_CRD","Dashboard vs SAP Result_CRD",
        "PSDD","Infor PSDD","Dashboard PSDD","Result_PSDD","Dashboard vs SAP Result_PSDD",
        "FCR Date","Dashboard FGR","Dashboard vs SAP Result FCR",
        "PODD","Infor PODD","Dashboard PODD","Result_PODD","Dashboard vs SAP Result_PODD",
        "PD","Infor PD","Dashboard PD","Result_PD","Dashboard vs SAP Result_PD",
        "PO Date","Actual PGI","Segment","S&P LPD","Currency"
    ]

    for col in final_order:
        if col not in df_out.columns:
            df_out[col] = pd.NA

    df_final = df_out.reindex(columns=final_order)
    return df_final

# -------------------------
# Run
# -------------------------
if run_button:
    if (pbi_file is None) or (sap_file is None):
        st.error("Upload kedua file: PBI Data dan SAP/INFOR report terlebih dahulu.")
    else:
        with st.spinner("Menjalankan pipeline..."):
            try:
                df_final = process(pbi_file, sap_file)
                st.success("Pipeline selesai — hasil tersedia di preview & download.")
            except Exception as e:
                st.exception(e)
                st.stop()

        # diagnostics & preview
        with st.expander("Diagnostik & ringkasan"):
            st.write("Rows:", df_final.shape[0], "Columns:", df_final.shape[1])
            st.write("Dtype sample:")
            for c in ["Customer PO item","Line Aggregator"]:
                if c in df_final.columns:
                    st.write(f"- {c}: dtype={df_final[c].dtype}; sample:", df_final[c].dropna().astype(str).head(5).tolist())
                else:
                    st.warning(f"Kolom '{c}' tidak ada di hasil!")

            dup_counts = None
            try:
                dup_counts = df_final["PO No.(Full)"].value_counts()
                st.write("Contoh PO multi-baris (top 10):")
                st.write(dup_counts.head(10))
            except Exception:
                pass

        if show_preview:
            st.subheader("Preview df_final (head)")
            st.dataframe(df_final.head(200), use_container_width=True)

        # prepare download as excel (in-memory)
        towrite = io.BytesIO()
        date_fmt_str = "mm/dd/yyyy"
        with pd.ExcelWriter(towrite, engine="xlsxwriter", datetime_format=date_fmt_str, date_format=date_fmt_str) as xw:
            df_final.to_excel(xw, sheet_name="Merged (Exploded)", index=False)
            ws = xw.sheets["Merged (Exploded)"]
            # apply date formatting to known date cols if present
            date_cols_exist = [c for c in [
                "Document Date","FPD","Infor FPD","Dashboard FPD",
                "LPD","Infor LPD","Dashboard LPD",
                "CRD","Infor CRD","Dashboard CRD",
                "PSDD","Infor PSDD","Dashboard PSDD",
                "FCR Date","Dashboard FGR",
                "PODD","Infor PODD","Dashboard PODD",
                "PD","Infor PD","Dashboard PD",
                "PO Date","Actual PGI"
            ] if c in df_final.columns]
            date_fmt = xw.book.add_format({"num_format": date_fmt_str})
            for c in date_cols_exist:
                col_idx = df_final.columns.get_loc(c)
                ws.set_column(col_idx, col_idx, 14, date_fmt)

            # autosize
            for i, col in enumerate(df_final.columns):
                sample_vals = df_final[col].astype(str).head(200).tolist()
                maxlen = max([len(str(col))] + [len(v) for v in sample_vals]) + 2
                ws.set_column(i, i, min(maxlen, 45))

        towrite.seek(0)
        now_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"PGD_sapinf_dashboard_exploded_{now_ts}.xlsx"
        st.download_button("Download hasil Excel", data=towrite, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.info("Upload file di sidebar lalu tekan tombol **Jalankan pipeline & buat file**.")
