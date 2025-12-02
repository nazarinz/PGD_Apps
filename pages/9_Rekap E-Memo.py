import streamlit as st
import os
import glob
import traceback
import shutil
from pathlib import Path
from datetime import datetime
from typing import List, Dict
import pandas as pd
import zipfile
import io
import tempfile

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

def extract_zip_to_temp(zip_file):
    """Extract uploaded ZIP to temporary directory"""
    temp_dir = tempfile.mkdtemp()
    
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)
    
    return temp_dir

def find_excel_files(root_dir):
    """Find all Excel/ODS files recursively"""
    exts = ["*.xls", "*.xlsx", "*.xlsm", "*.xlsb", "*.ods"]
    files = []
    for ext in exts:
        files += glob.glob(os.path.join(root_dir, "**", ext), recursive=True)
    return sorted(files)

def process_folder(root_dir):
    """Process all Excel files in folder"""
    files = find_excel_files(root_dir)
    rows, skipped, no_keys = [], [], []
    today_date = pd.Timestamp.today().normalize()
    total = len(files)
    
    if total == 0:
        return pd.DataFrame(columns=OUTPUT_COLS), [], [], []
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    file_list = []
    for idx, file_path in enumerate(files, 1):
        rel_path = os.path.relpath(file_path, root_dir)
        file_list.append(rel_path)
        status_text.text(f"Processing {idx}/{total}: {rel_path}")
        progress_bar.progress(idx / total)
        
        try:
            row = parse_excel_file(file_path)
            all_empty = all((not str(row[c]).strip()) for c in TARGET_HEADERS)
            if all_empty:
                no_keys.append(rel_path)
                continue
            row["_source_file"] = rel_path
            row["Tanggal Rekap"] = today_date
            rows.append(row)
        except Exception as e:
            tb = traceback.format_exc()
            skipped.append((rel_path, str(e)))
    
    status_text.text("✅ Processing complete!")
    
    if rows:
        df_out = pd.DataFrame(rows).reindex(columns=OUTPUT_COLS)
    else:
        df_out = pd.DataFrame(columns=OUTPUT_COLS)
    
    return df_out, no_keys, skipped, file_list

# ---------------- Streamlit UI ----------------
st.set_page_config(page_title="Excel/ODS Parser", page_icon="📊", layout="wide")

st.title("📊 Excel/ODS Parser untuk E-Memo")
st.markdown("Upload folder (dalam format ZIP) yang berisi file Excel/ODS untuk diparse secara otomatis")

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
    
    **Cara Upload Folder:**
    1. Compress folder Anda menjadi ZIP
    2. Upload file ZIP
    3. Semua file Excel/ODS di dalam folder (termasuk subfolder) akan diproses otomatis
    
    **Target Headers:**
    """)
    with st.expander("Lihat semua headers"):
        for header in TARGET_HEADERS:
            st.text(f"• {header}")

# Instructions
st.info("💡 **Cara Upload Folder:** Compress folder Anda menjadi file ZIP, kemudian upload file ZIP tersebut di bawah ini.")

# File uploader for ZIP
uploaded_zip = st.file_uploader(
    "Upload Folder (dalam format ZIP)",
    type=["zip"],
    help="Upload file ZIP yang berisi folder dengan file Excel/ODS"
)

if uploaded_zip:
    st.success(f"✅ File ZIP berhasil diupload: {uploaded_zip.name}")
    
    if st.button("🚀 Mulai Processing", type="primary"):
        temp_dir = None
        try:
            with st.spinner("Extracting ZIP file..."):
                temp_dir = extract_zip_to_temp(uploaded_zip)
                st.success("✅ ZIP file berhasil di-extract")
            
            with st.spinner("Scanning dan memproses file..."):
                df_result, no_keys, skipped, file_list = process_folder(temp_dir)
                
                # Display results
                st.header("📈 Hasil Rekap")
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total File Ditemukan", len(file_list))
                with col2:
                    st.metric("Berhasil Diparse", len(df_result))
                with col3:
                    st.metric("Label Tidak Ketemu", len(no_keys))
                with col4:
                    st.metric("Error", len(skipped))
                
                # Show file list
                with st.expander(f"📁 Daftar File yang Ditemukan ({len(file_list)})"):
                    for f in file_list:
                        st.text(f"• {f}")
                
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
                        df_result.to_excel(writer, index=False, sheet_name='Data')
                        
                        # Add summary sheet
                        summary_data = {
                            'Metric': ['Total File Ditemukan', 'Berhasil Diparse', 'Label Tidak Ketemu', 'Error'],
                            'Count': [len(file_list), len(df_result), len(no_keys), len(skipped)]
                        }
                        pd.DataFrame(summary_data).to_excel(writer, index=False, sheet_name='Summary')
                        
                        # Add skipped files sheet if any
                        if skipped:
                            pd.DataFrame(skipped, columns=['File', 'Error']).to_excel(writer, index=False, sheet_name='Errors')
                        
                        # Add no_keys files sheet if any
                        if no_keys:
                            pd.DataFrame({'File': no_keys}).to_excel(writer, index=False, sheet_name='No Keys')
                    
                    output.seek(0)
                    
                    st.download_button(
                        label="📥 Download Hasil (Excel)",
                        data=output,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
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
                            with st.container():
                                st.code(error, language=None)
                        if len(skipped) > 10:
                            st.text(f"... dan {len(skipped) - 10} error lainnya")
                
        except Exception as e:
            st.error(f"❌ Terjadi error: {str(e)}")
            st.code(traceback.format_exc())
        
        finally:
            # Cleanup temp directory
            if temp_dir and os.path.exists(temp_dir):
                try:
                    shutil.rmtree(temp_dir)
                except:
                    pass

else:
    st.info("👆 Silakan upload file ZIP yang berisi folder dengan file Excel/ODS")
    
    # Show example
    with st.expander("💡 Panduan Lengkap"):
        st.markdown("""
        ### Langkah-langkah:
        
        1. **Persiapan Folder**
           - Pastikan semua file Excel/ODS ada dalam satu folder
           - File bisa tersebar di berbagai subfolder
        
        2. **Compress ke ZIP**
           - **Windows:** Klik kanan folder → Send to → Compressed (zipped) folder
           - **Mac:** Klik kanan folder → Compress
           - **Linux:** Klik kanan folder → Compress
        
        3. **Upload ZIP**
           - Click tombol **"Browse files"** di atas
           - Pilih file ZIP yang sudah dibuat
        
        4. **Processing**
           - Click tombol **"Mulai Processing"**
           - Tunggu hingga semua file diproses
        
        5. **Download Hasil**
           - Download hasil dalam format Excel
           - File hasil berisi beberapa sheet:
             - **Data:** Hasil parse semua file
             - **Summary:** Ringkasan statistik
             - **Errors:** Daftar file yang error (jika ada)
             - **No Keys:** File yang tidak ketemu labelnya (jika ada)
        
        ### Contoh Struktur Folder:
        ```
        Rekap E-Memo/
        ├── Januari/
        │   ├── file1.xlsx
        │   └── file2.xls
        ├── Februari/
        │   ├── file3.xlsx
        │   └── subfolder/
        │       └── file4.ods
        └── file5.xlsm
        ```
        
        **Semua file akan diproses secara otomatis, termasuk yang ada di subfolder!**
        """)

# Footer
st.markdown("---")
st.markdown("Parser untuk E-Memo Rekap")
