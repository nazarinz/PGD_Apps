import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime
from io import BytesIO

# ========================================
# Page Configuration
# ========================================
st.set_page_config(
    page_title="PGD Data Analysis",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ========================================
# Custom CSS
# ========================================
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f2937;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #6b7280;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d1fae5;
        border: 1px solid #10b981;
        margin: 1rem 0;
    }
    .info-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #dbeafe;
        border: 1px solid #3b82f6;
        margin: 1rem 0;
    }
    .warning-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #fef3c7;
        border: 1px solid #f59e0b;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# ========================================
# Helper Functions
# ========================================
def _clean_cols(pdf: pd.DataFrame) -> pd.DataFrame:
    """Rapikan nama kolom: ganti spasi ganda & trim pinggir."""
    pdf = pdf.copy()
    pdf.columns = (
        pdf.columns
        .astype(str)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    return pdf

def _ref_qty_sap_row(row: pd.Series):
    """Ambil kuantitas referensi per-baris SAP: Infor Quantity > Quantity."""
    iq = pd.to_numeric(row.get("Infor Quantity"), errors="coerce")
    q = pd.to_numeric(row.get("Quantity"), errors="coerce")
    return iq if pd.notna(iq) else q

def _num(s):
    return pd.to_numeric(s, errors="coerce")

def _to_dt(s):
    return pd.to_datetime(s, errors="coerce")

def compare_numeric(df, left, right):
    a, b = _num(df.get(left)), _num(df.get(right))
    return (a.notna() & b.notna() & (a == b))

def compare_date(df, left, right):
    a, b = _to_dt(df.get(left)), _to_dt(df.get(right))
    return (a.notna() & b.notna() & (a.dt.date == b.dt.date))

def _build_pairs_for_po(df_sap_po: pd.DataFrame, df_dash_po: pd.DataFrame):
    """Build 1:1 pairing berdasarkan nearest quantity."""
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
            candidates.append((dist, -(dq if pd.notna(dq) else -np.inf), pos_map[i], i, j))
    
    candidates.sort(key=lambda x: (x[0], x[1], x[2]))
    
    used_sap, used_dash, chosen = set(), set(), []
    for dist, _negdq, _ord, i, j in candidates:
        if i in used_sap or j in used_dash:
            continue
        chosen.append((i, j))
        used_sap.add(i)
        used_dash.add(j)
    
    for i in sap_idx:
        if i not in used_sap:
            chosen.append((i, None))
    for j in dash_idx:
        if j not in used_dash:
            chosen.append((None, j))
    
    return chosen

def _norm_str_col(s):
    """Normalize string column."""
    return s.astype("string").str.strip().str.upper()

# ========================================
# Main Processing Function
# ========================================
def process_pgd_data(pbi_file, sapinf_file, progress_bar, status_text):
    """Main processing function untuk PGD data analysis."""
    
    try:
        # Step 1: Load Data
        status_text.text("📂 Loading PBI Data...")
        progress_bar.progress(10)
        df_db = pd.read_excel(pbi_file, engine="openpyxl")
        df_db = _clean_cols(df_db)
        
        status_text.text("📂 Loading SAP/INFOR report...")
        progress_bar.progress(20)
        df_sapinf = pd.read_excel(sapinf_file, engine="openpyxl")
        df_sapinf = _clean_cols(df_sapinf)
        
        # Step 2: Clean duplicate columns
        status_text.text("🧹 Cleaning duplicate columns...")
        progress_bar.progress(30)
        
        PK = "PO Number (GPS)"
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
        
        dashboard_alias = {
            "Elevated check": "Elevated Check",
            "Elevated_Check": "Elevated Check",
            "Responsiveness Status": "Responsiveness",
            "Responsiveness_": "Responsiveness",
            "PO Status (Dashboard)": "PO Status",
        }
        df.rename(columns={k:v for k,v in dashboard_alias.items() if k in df.columns}, inplace=True)
        
        if PK not in df.columns:
            raise KeyError(f"Kolom PK '{PK}' tidak ada di PBI Data.")
        
        sapinf_key_candidates = ["PO Number (GPS)", "PO No.(Full)", "PO No.", "PO Number"]
        join_left_key = next((k for k in sapinf_key_candidates if k in df_sapinf.columns), None)
        
        if join_left_key is None:
            raise KeyError(f"Tidak menemukan key join di SAP/INFOR file.")
        
        df_sapinf_join = df_sapinf.copy()
        if join_left_key != PK:
            df_sapinf_join[PK] = df_sapinf_join[join_left_key]
        
        # Step 3: Prepare Dashboard columns
        status_text.text("📊 Preparing dashboard columns...")
        progress_bar.progress(40)
        
        dash_map_src_to_new = {
            "Order ALL": "Dashboard Quantity",
            "MDP Status Adjusted": "Dashboard MDP Status Adjusted",
            "SDP Status Adjusted": "Dashboard SDP Status Adjusted",
            "FPD_": "Dashboard FPD",
            "LPD_": "Dashboard LPD",
            "CRD_": "Dashboard CRD",
            "PSDD_": "Dashboard PSDD",
            "Planned Date_": "Dashboard PD",
            "PODD_": "Dashboard PODD",
            "FGR Document Date_": "Dashboard FGR",
        }
        
        present_src = [src for src in dash_map_src_to_new if src in df.columns]
        passthrough_cols = ["Elevated Check", "Responsiveness", "PO Status"]
        present_passthrough = [c for c in passthrough_cols if c in df.columns]
        
        keep_cols = [PK] + present_src + present_passthrough
        df_dash_keep = df[keep_cols].copy(deep=True)
        df_dash_keep = df_dash_keep.rename(columns={src: dash_map_src_to_new[src] for src in present_src})
        
        for src, newc in dash_map_src_to_new.items():
            if newc not in df_dash_keep.columns:
                df_dash_keep[newc] = pd.NA
        
        for c in passthrough_cols:
            if c not in df_dash_keep.columns:
                df_dash_keep[c] = pd.NA
        
        # Step 4: Pairing 1:1
        status_text.text("🔗 Performing 1:1 nearest quantity pairing...")
        progress_bar.progress(50)
        
        out_rows = []
        for po, df_sap_po in df_sapinf_join.groupby(PK, dropna=False):
            df_dash_po = df_dash_keep[df_dash_keep[PK] == po]
            pairs = _build_pairs_for_po(df_sap_po, df_dash_po)
            
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
        
        # Clean up columns
        for c in ["Elevated Check", "Responsiveness", "PO Status"]:
            if c in df_out.columns:
                df_out[c] = df_out[c].astype("string").str.strip()
        
        for c in ["Quantity", "Infor Quantity", "Dashboard Quantity"]:
            if c in df_out.columns:
                df_out[c] = pd.to_numeric(df_out[c], errors="coerce")
        
        # Step 5: Create comparison columns
        status_text.text("📋 Creating comparison columns...")
        progress_bar.progress(60)
        
        if ("Quantity" in df_out.columns) and ("Infor Quantity" in df_out.columns):
            df_out["Result_Quantity"] = compare_numeric(df_out, "Quantity", "Infor Quantity")
        else:
            df_out["Result_Quantity"] = pd.NA
        
        if ("Quantity" in df_out.columns) and ("Dashboard Quantity" in df_out.columns):
            df_out["Dashboard vs SAP Result_Quantity"] = compare_numeric(df_out, "Quantity", "Dashboard Quantity")
        else:
            df_out["Dashboard vs SAP Result_Quantity"] = pd.NA
        
        date_compare_pairs = [
            ("FPD", "Infor FPD", "Dashboard FPD", "Result_FPD", "Dashboard vs SAP Result_FPD"),
            ("LPD", "Infor LPD", "Dashboard LPD", "Result_LPD", "Dashboard vs SAP Result_LPD"),
            ("CRD", "Infor CRD", "Dashboard CRD", "Result_CRD", "Dashboard vs SAP Result_CRD"),
            ("PSDD", "Infor PSDD", "Dashboard PSDD", "Result_PSDD", "Dashboard vs SAP Result_PSDD"),
            ("PODD", "Infor PODD", "Dashboard PODD", "Result_PODD", "Dashboard vs SAP Result_PODD"),
            ("PD", "Infor PD", "Dashboard PD", "Result_PD", "Dashboard vs SAP Result_PD"),
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
        
        # Step 6: MDP/SDP Delay Qty & GAP
        status_text.text("📈 Calculating MDP/SDP Delay Quantities...")
        progress_bar.progress(70)
        
        for need in ["MDP","SDP","Dashboard MDP Status Adjusted","Dashboard SDP Status Adjusted",
                     "Quantity","Dashboard Quantity"]:
            if need not in df_out.columns:
                df_out[need] = pd.NA
        
        qty_sap_num = pd.to_numeric(df_out["Quantity"], errors="coerce").fillna(0)
        qty_dash_num = pd.to_numeric(df_out["Dashboard Quantity"], errors="coerce").fillna(0)
        
        mdp_status = _norm_str_col(df_out["MDP"])
        dash_mdp_stat = _norm_str_col(df_out["Dashboard MDP Status Adjusted"])
        sdp_status = _norm_str_col(df_out["SDP"])
        dash_sdp_stat = _norm_str_col(df_out["Dashboard SDP Status Adjusted"])
        
        mask_mdp_fail = mdp_status.eq("FAIL").fillna(False).to_numpy()
        mask_dash_mdp_delay = dash_mdp_stat.eq("DELAY").fillna(False).to_numpy()
        mask_sdp_fail = sdp_status.eq("FAIL").fillna(False).to_numpy()
        mask_dash_sdp_delay = dash_sdp_stat.eq("DELAY").fillna(False).to_numpy()
        
        df_out["MDP Delay Qty"] = np.where(mask_mdp_fail, qty_sap_num, 0)
        df_out["Dashboard MDP Delay Qty"] = np.where(mask_dash_mdp_delay, qty_dash_num, 0)
        df_out["GAP MDP"] = df_out["Dashboard MDP Delay Qty"] - df_out["MDP Delay Qty"]
        
        df_out["SDP Delay Qty"] = np.where(mask_sdp_fail, qty_sap_num, 0)
        df_out["Dashboard SDP Delay Qty"] = np.where(mask_dash_sdp_delay, qty_dash_num, 0)
        df_out["GAP SDP"] = df_out["Dashboard SDP Delay Qty"] - df_out["SDP Delay Qty"]
        
        for c in ["MDP Delay Qty","Dashboard MDP Delay Qty","GAP MDP",
                  "SDP Delay Qty","Dashboard SDP Delay Qty","GAP SDP"]:
            df_out[c] = pd.to_numeric(df_out[c], errors="coerce").round(0).astype("Int64")
        
        gap_mdp_num = pd.to_numeric(df_out["GAP MDP"], errors="coerce")
        gap_sdp_num = pd.to_numeric(df_out["GAP SDP"], errors="coerce")
        
        df_out["GAP MDP (NZ)"] = gap_mdp_num.where(gap_mdp_num.ne(0)).round(0).astype("Int64")
        df_out["GAP SDP (NZ)"] = gap_sdp_num.where(gap_sdp_num.ne(0)).round(0).astype("Int64")
        
        df_out["Has MDP Gap"] = gap_mdp_num.fillna(0).ne(0)
        df_out["Has SDP Gap"] = gap_sdp_num.fillna(0).ne(0)
        
        # Step 7: Normalize dates & CRD Monthly
        status_text.text("📅 Processing dates and CRD Monthly...")
        progress_bar.progress(80)
        
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
        
        # Step 8: Final column ordering
        status_text.text("🎯 Finalizing output...")
        progress_bar.progress(90)
        
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
        
        progress_bar.progress(100)
        status_text.text("✅ Processing complete!")
        
        return df_final, None
        
    except Exception as e:
        return None, str(e)

def convert_df_to_excel(df):
    """Convert dataframe to Excel with proper formatting."""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Merged (Exploded)', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Merged (Exploded)']
        
        # Date format
        date_format = workbook.add_format({'num_format': 'mm/dd/yyyy'})
        
        # Apply date format to date columns
        date_cols = [
            "Document Date","FPD","Infor FPD","Dashboard FPD",
            "LPD","Infor LPD","Dashboard LPD",
            "CRD","Infor CRD","Dashboard CRD",
            "PSDD","Infor PSDD","Dashboard PSDD",
            "FCR Date","Dashboard FGR",
            "PODD","Infor PODD","Dashboard PODD",
            "PD","Infor PD","Dashboard PD",
            "PO Date","Actual PGI"
        ]
        
        for col in date_cols:
            if col in df.columns:
                col_idx = df.columns.get_loc(col)
                worksheet.set_column(col_idx, col_idx, 14, date_format)
        
        # Auto-size columns
        for i, col in enumerate(df.columns):
            max_len = max(
                df[col].astype(str).apply(len).max(),
                len(str(col))
            ) + 2
            worksheet.set_column(i, i, min(max_len, 45))
    
    output.seek(0)
    return output

# ========================================
# Streamlit App
# ========================================
def main():
    # Header
    st.markdown('<p class="main-header">📊 PGD Data Analysis</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Merge per-PO (Exploded) + Dashboard vs SAP Compare (No Aggregation)</p>', unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.header("ℹ️ About")
        st.info("""
        **Version:** 2025-10-21
        
        **Features:**
        - 1:1 Nearest quantity pairing per PO
        - Customer PO item & Line Aggregator from SAP
        - MDP/SDP Delay Qty calculations
        - GAP analysis (including NZ columns)
        - CRD Monthly added
        - All comparison columns
        """)
        
        st.header("📋 Processing Details")
        st.markdown("""
        **Pairing Method:** 1:1 nearest quantity matching per PO
        
        **Key Fields:** Customer PO item & Line Aggregator taken from SAP (raw)
        
        **Calculations:**
        - MDP/SDP Delay Qty
        - Dashboard Delay Qty
        - GAP, GAP (NZ)
        - Has Gap indicators
        
        **Date Handling:** CRD Monthly added before CRD, all dates normalized
        
        **Comparisons:** Dashboard vs SAP for Quantity, FPD, LPD, CRD, PSDD, PODD, PD, FCR
        """)
    
    # Main content
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📁 Upload PBI Data File")
        pbi_file = st.file_uploader(
            "Dashboard (PBI) Excel file",
            type=['xlsx', 'xls'],
            key='pbi',
            help="Upload your PBI Data Excel file"
        )
        if pbi_file:
            st.success(f"✓ {pbi_file.name}")
    
    with col2:
        st.subheader("📁 Upload SAP/INFOR File")
        sapinf_file = st.file_uploader(
            "Comparison Tracking Report",
            type=['xlsx', 'xls'],
            key='sapinf',
            help="Upload your SAP/INFOR Excel file"
        )
        if sapinf_file:
            st.success(f"✓ {sapinf_file.name}")
    
    st.markdown("---")
    
    # Process button
    if st.button("🚀 Process Files", type="primary", use_container_width=True):
        if not pbi_file or not sapinf_file:
            st.error("❌ Please upload both files before processing.")
        else:
            # Progress indicators
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Process files
            df_result, error = process_pgd_data(pbi_file, sapinf_file, progress_bar, status_text)
            
            if error:
                st.error(f"❌ Error processing files: {error}")
            else:
                st.markdown('<div class="success-box">', unsafe_allow_html=True)
                st.success("✅ Processing completed successfully!")
                
                # Display summary
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Rows", f"{len(df_result):,}")
                with col2:
                    st.metric("Total Columns", len(df_result.columns))
                with col3:
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    st.metric("Processed At", timestamp)
                
                st.markdown('</div>', unsafe_allow_html=True)
                
                # Preview
                st.subheader("📊 Data Preview")
                st.dataframe(df_result.head(100), use_container_width=True, height=400)
                
                # Download button
                st.subheader("💾 Download Result")
                
                excel_file = convert_df_to_excel(df_result)
                timestamp_file = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"PGD_sapinf_dashboard_exploded_{timestamp_file}.xlsx"
                
                st.download_button(
                    label="📥 Download Excel File",
                    data=excel_file,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                # Additional info
                st.markdown('<div class="info-box">', unsafe_allow_html=True)
                st.markdown("""
                **✅ Features included in this export:**
                - 1:1 Nearest quantity pairing per PO
                - Customer PO item & Line Aggregator from SAP
                - MDP/SDP Delay Qty calculations
                - GAP analysis (including NZ columns)
                - CRD Monthly added before CRD
                - All comparison columns (SAP vs Dashboard)
                - Properly formatted dates (mm/dd/yyyy)
                """)
                st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
