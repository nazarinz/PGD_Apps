import streamlit as st
import os
import glob
import traceback
from pathlib import Path
from datetime import datetime
from typing import List, Dict
import pandas as pd
import zipfile
import io

# ---------------- CONFIG ----------------
TARGET_HEADERS = [
    "Last","Model Name","Mat.Way ID","Model Number","Collar Lat. Height","Bottom ID",
    "Sample Size","Spec #","Developer/Pattern","Article No.","Heel Height","Upper ID",
    "ETC","Quantity","Stage/Purpose","Gender/Age Group","Sign CWA","Pro-time",
    "Cutting Die","Spec.# Chg","Season/RI Date","Inner Length","Type(Model/Article)",
    "Factory","Project Manager","LO","Region/Category","Level","Material Way",
    "Leadtime","Size Run"
]
KEY_SET = set(TARGET_HEADERS)
ALIAS = {
    "Spec#": "Spec #", "Spec No.": "Spec #",
    "ArticleNo.": "Article No.", "ArticleNo": "Article No.",
    "ModelName": "Model Name", "Mat Way ID": "Mat.Way ID",
    "MatWayID": "Mat.Way ID", "Gender": "Gender/Age Group",
    "Gender/Age": "Gender/Age Group", "Prod Manager": "Project Manager",
    "Region": "Region/Category", "Category": "Region/Category",
}
OUTPUT_COLS = [
    "_source_file","Tanggal Rekap","Model Number","Article No.","Last","Model Name",
    "Mat.Way ID","Collar Lat. Height","Bottom ID","Sample Size","Spec #",
    "Developer/Pattern","Heel Height","Upper ID","ETC","Quantity","Stage/Purpose",
    "Gender/Age Group","Sign CWA","Pro-time","Cutting Die","Spec.# Chg",
    "Season/RI Date","Inner Length","Type(Model/Article)","Factory","Project Manager",
    "LO","Region/Category","Level","Material Way","Leadtime","Size Run"
]

# ---------------- IO helpers ----------------
def safe_read_any(path: str) -> pd.DataFrame:
    ext = Path(path).suffix.lower()
    if ext in [".xlsx", ".xlsm"]:
        return pd.read_excel(path, header=None, engine="openpyxl")
    elif ext == ".xls":
        return pd.read_excel(path, header=None, engine="xlrd")
    elif ext == ".xlsb":
        return pd.read_excel(path, header=None, engine="pyxlsb")
    elif ext == ".ods":
        return pd.read_excel(path, header=None, engine="odf")
    else:
        return pd.read_excel(path, header=None)

def norm_cell_value(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    if not s:
        return None
    if s in ALIAS:
        return ALIAS[s]
    return s

def row_to_mapping(cells: List) -> Dict[str, str]:
    norm = [norm_cell_value(c) for c in cells]
    key_positions = [j for j, s in enumerate(norm) if s in KEY_SET]
    if not key_positions:
        key_positions = [j for j, s in enumerate(norm) if s and any(k.lower() in s.lower() for k in KEY_SET)]
    if not key_positions:
        return {}
    mapping = {}
    for idx, key_idx in enumerate(key_positions):
        next_bound = len(norm)
        if idx + 1 < len(key_positions):
            next_bound = key_positions[idx+1]
        key = norm[key_idx]
        val = None
        for j in range(key_idx+1, next_bound):
            s = norm[j]
            if s and s not in KEY_SET:
                val = s
                break
        if not val:
            raw = cells[key_idx]
            if isinstance(raw, str) and ":" in raw:
                parts = raw.split(":", 1)
                if parts[0].strip() in KEY_SET and parts[1].strip():
                    val = parts[1].strip()
        if val and key not in mapping:
            mapping[key] = val
    return mapping

def parse_excel_file(path: str) -> Dict[str,str]:
    df = safe_read_any(path)
    collected = {}
    for i in range(min(len(df), 400)):
        try:
            m = row_to_mapping(df.iloc[i,:].tolist())
        except Exception:
            m = {}
        for k, v in m.items():
            if k not in collected and v is not None:
                collected[k] = v
        if set(collected.keys()) >= KEY_SET:
            break
    return {col: collected.get(col, "") for col in TARGET_HEADERS}

def process_files(files, temp_dir):
    rows, skipped, no_keys = [], [], []
    today_date = pd.Timestamp.today().normalize()
    total = len(files)
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, uploaded_file in enumerate(files, 1):
        status_text.text(f"Processing {idx}/{total}: {uploaded_file.name}")
        progress_bar.progress(idx / total)
        
        # Save uploaded file temporarily
        temp_path = os.path.join(temp_dir, uploaded_file.name)
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        try:
            row = parse_excel_file(temp_path)
            all_empty = all((not str(row[c]).strip()) for c in TARGET_HEADERS)
            if all_empty:
                no_keys.append(uploaded_file.name)
                continue
            row["_source_file"] = uploaded_file.name
            row["Tanggal Rekap"] = today_date
            rows.append(row)
        except Exception as e:
            skipped.append((uploaded_file.name, str(e)))
    
    status_text.text("Processing complete!")
    
    if rows:
        df_out = pd.DataFrame(rows).reindex(columns=OUTPUT_COLS)
    else:
        df_out = pd.DataFrame(columns=OUTPUT_COLS)
    
    return df_out, no_keys, skipped

# ---------------- Streamlit UI ----------------
st.set_page_config(page_title="Excel/ODS Parser", page_icon="📊", layout="wide")

st.title("📊 Excel/ODS Parser untuk E-Memo")
st.markdown("Upload file Excel (.xls, .xlsx, .xlsm, .xlsb) atau ODS untuk diparse secara otomatis")

# Sidebar info
with st.sidebar:
    st.header("ℹ️ Informasi")
    st.markdown("""
    **Format yang didukung:**
    - .xls
    - .xlsx
    - .xlsm
    - .xlsb
    - .ods
    
    **Target Headers:**
    """)
    with st.expander("Lihat semua headers"):
        for header in TARGET_HEADERS:
            st.text(f"• {header}")

# File uploader
uploaded_files = st.file_uploader(
    "Upload file Excel/ODS",
    type=["xls", "xlsx", "xlsm", "xlsb", "ods"],
    accept_multiple_files=True,
    help="Anda dapat upload multiple files sekaligus"
)

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} file berhasil diupload")
    
    # Create temporary directory
    temp_dir = "temp_uploads"
    os.makedirs(temp_dir, exist_ok=True)
    
    if st.button("🚀 Mulai Processing", type="primary"):
        with st.spinner("Sedang memproses file..."):
            try:
                df_result, no_keys, skipped = process_files(uploaded_files, temp_dir)
                
                # Display results
                st.header("📈 Hasil Rekap")
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total File", len(uploaded_files))
                with col2:
                    st.metric("Berhasil Diparse", len(df_result))
                with col3:
                    st.metric("Label Tidak Ketemu", len(no_keys))
                with col4:
                    st.metric("Error", len(skipped))
                
                # Show parsed data
                if not df_result.empty:
                    st.subheader("📋 Data Hasil Parse")
                    st.dataframe(df_result, use_container_width=True, height=400)
                    
                    # Download button
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_filename = f"Result_{timestamp}.xlsx"
                    
                    # Create Excel file in memory
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_result.to_excel(writer, index=False)
                    output.seek(0)
                    
                    st.download_button(
                        label="📥 Download Hasil (Excel)",
                        data=output,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("⚠️ Tidak ada data yang berhasil diparse")
                
                # Show no_keys files
                if no_keys:
                    with st.expander(f"⚠️ File dengan Label Tidak Ketemu ({len(no_keys)})"):
                        for f in no_keys:
                            st.text(f"• {f}")
                
                # Show errors
                if skipped:
                    with st.expander(f"❌ File dengan Error ({len(skipped)})"):
                        for fname, error in skipped[:10]:
                            st.text(f"• {fname}")
                            st.code(error, language=None)
                        if len(skipped) > 10:
                            st.text(f"... dan {len(skipped) - 10} error lainnya")
                
            except Exception as e:
                st.error(f"❌ Terjadi error: {str(e)}")
                st.code(traceback.format_exc())
            
            finally:
                # Cleanup temp directory
                try:
                    import shutil
                    shutil.rmtree(temp_dir)
                except:
                    pass

else:
    st.info("👆 Silakan upload file untuk memulai")
    
    # Show example
    with st.expander("💡 Cara Penggunaan"):
        st.markdown("""
        1. Click tombol **"Browse files"** di atas
        2. Pilih satu atau lebih file Excel/ODS
        3. Click tombol **"Mulai Processing"**
        4. Tunggu hingga proses selesai
        5. Download hasil dalam format Excel
        
        **Catatan:** 
        - File akan diparse untuk mencari header yang sesuai
        - Hasil akan mencakup informasi dari semua file yang berhasil diparse
        - File yang gagal akan ditampilkan dalam daftar error
        """)

# Footer
st.markdown("---")
st.markdown("Made with ❤️ using Streamlit | Parser untuk E-Memo Rekap")
