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

def extract_zip_to_temp(zip_file, zip_name):
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

def process_multiple_zips(uploaded_zips):
    """Process multiple ZIP files"""
    all_rows = []
    all_skipped = []
    all_no_keys = []
    all_file_list = []
    temp_dirs = []
    
    today_date = pd.Timestamp.today().normalize()
    
    # Extract all ZIPs first
    st.info(f"📦 Extracting {len(uploaded_zips)} ZIP file(s)...")
    extraction_progress = st.progress(0)
    
    for idx, zip_file in enumerate(uploaded_zips):
        try:
            temp_dir = extract_zip_to_temp(zip_file, zip_file.name)
            temp_dirs.append((temp_dir, zip_file.name))
            extraction_progress.progress((idx + 1) / len(uploaded_zips))
        except Exception as e:
            st.error(f"❌ Error extracting {zip_file.name}: {str(e)}")
            continue
    
    st.success(f"✅ {len(temp_dirs)} ZIP file(s) berhasil di-extract")
    
    # Process all files
    total_files = 0
    for temp_dir, zip_name in temp_dirs:
        files = find_excel_files(temp_dir)
        total_files += len(files)
    
    if total_files == 0:
        return pd.DataFrame(columns=OUTPUT_COLS), [], [], [], temp_dirs
    
    st.info(f"🔍 Total {total_files} file Excel/ODS ditemukan")
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    current_file_idx = 0
    
    # Process each ZIP
    for temp_dir, zip_name in temp_dirs:
        files = find_excel_files(temp_dir)
        
        for file_path in files:
            current_file_idx += 1
            rel_path = os.path.relpath(file_path, temp_dir)
            # Add ZIP name prefix to identify source
            full_path = f"[{zip_name}] {rel_path}"
            all_file_list.append(full_path)
            
            status_text.text(f"Processing {current_file_idx}/{total_files}: {full_path}")
            progress_bar.progress(current_file_idx / total_files)
            
            try:
                row = parse_excel_file(file_path)
                all_empty = all((not str(row[c]).strip()) for c in TARGET_HEADERS)
                if all_empty:
                    all_no_keys.append(full_path)
                    continue
                row["_source_file"] = full_path
                row["Tanggal Rekap"] = today_date
                all_rows.append(row)
            except Exception as e:
                tb = traceback.format_exc()
                all_skipped.append((full_path, str(e)))
    
    status_text.text("✅ Processing complete!")
    
    if all_rows:
        df_out = pd.DataFrame(all_rows).reindex(columns=OUTPUT_COLS)
    else:
        df_out = pd.DataFrame(columns=OUTPUT_COLS)
    
    return df_out, all_no_keys, all_skipped, all_file_list, temp_dirs

# ---------------- Streamlit UI ----------------
st.set_page_config(page_title="Excel/ODS Parser", page_icon="📊", layout="wide")

st.title("📊 Excel/ODS Parser untuk E-Memo")
st.markdown("Upload **multiple ZIP files** yang berisi file Excel/ODS untuk diparse secara otomatis")

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
    
    **Cara Upload Multiple ZIP:**
    1. Compress setiap folder menjadi ZIP
    2. Upload semua file ZIP sekaligus
    3. Semua file Excel/ODS akan diproses otomatis
    
    **Keuntungan Multiple ZIP:**
    - Upload beberapa folder sekaligus
    - Source file diberi label nama ZIP
    - Lebih cepat dan efisien
    
    **Target Headers:**
    """)
    with st.expander("Lihat semua headers"):
        for header in TARGET_HEADERS:
            st.text(f"• {header}")

# Instructions
st.info("💡 **Cara Upload:** Compress setiap folder menjadi file ZIP, kemudian upload **semua file ZIP sekaligus** di bawah ini.")

# File uploader for multiple ZIPs
uploaded_zips = st.file_uploader(
    "Upload Multiple ZIP Files",
    type=["zip"],
    accept_multiple_files=True,
    help="Anda bisa upload beberapa file ZIP sekaligus. Setiap ZIP bisa berisi folder dengan file Excel/ODS"
)

if uploaded_zips:
    st.success(f"✅ {len(uploaded_zips)} file ZIP berhasil diupload")
    
    # Show uploaded ZIP files
    with st.expander("📦 Daftar ZIP yang Diupload"):
        for idx, zip_file in enumerate(uploaded_zips, 1):
            st.text(f"{idx}. {zip_file.name} ({zip_file.size / 1024:.1f} KB)")
    
    if st.button("🚀 Mulai Processing Semua ZIP", type="primary"):
        temp_dirs = []
        try:
            with st.spinner("Processing multiple ZIP files..."):
                df_result, no_keys, skipped, file_list, temp_dirs = process_multiple_zips(uploaded_zips)
                
                # Display results
                st.header("📈 Hasil Rekap")
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total ZIP Diupload", len(uploaded_zips))
                with col2:
                    st.metric("Total File Ditemukan", len(file_list))
                with col3:
                    st.metric("Berhasil Diparse", len(df_result))
                with col4:
                    st.metric("Error + No Keys", len(skipped) + len(no_keys))
                
                # Show file list grouped by ZIP
                with st.expander(f"📁 Daftar Semua File yang Ditemukan ({len(file_list)})"):
                    current_zip = None
                    for f in file_list:
                        # Extract ZIP name from path
                        if f.startswith("["):
                            zip_name = f.split("]")[0][1:]
                            if zip_name != current_zip:
                                current_zip = zip_name
                                st.markdown(f"**📦 {zip_name}**")
                        st.text(f"  • {f}")
                
                # Show parsed data
                if not df_result.empty:
                    st.subheader("📋 Data Hasil Parse")
                    
                    # Add filter by ZIP
                    all_zips = ["Semua ZIP"] + [f"[{z.name}]" for _, z in zip(temp_dirs, uploaded_zips)]
                    selected_zip = st.selectbox("Filter berdasarkan ZIP:", all_zips)
                    
                    if selected_zip == "Semua ZIP":
                        filtered_df = df_result
                    else:
                        filtered_df = df_result[df_result["_source_file"].str.startswith(selected_zip)]
                    
                    st.dataframe(filtered_df, use_container_width=True, height=400)
                    st.info(f"Menampilkan {len(filtered_df)} dari {len(df_result)} data")
                    
                    # Download button
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_filename = f"Result_{timestamp}.xlsx"
                    
                    # Create Excel file in memory
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_result.to_excel(writer, index=False, sheet_name='Data')
                        
                        # Add summary sheet
                        summary_data = {
                            'Metric': [
                                'Total ZIP Diupload',
                                'Total File Ditemukan', 
                                'Berhasil Diparse', 
                                'Label Tidak Ketemu', 
                                'Error'
                            ],
                            'Count': [
                                len(uploaded_zips),
                                len(file_list), 
                                len(df_result), 
                                len(no_keys), 
                                len(skipped)
                            ]
                        }
                        pd.DataFrame(summary_data).to_excel(writer, index=False, sheet_name='Summary')
                        
                        # Add ZIP list
                        zip_data = {
                            'ZIP File': [z.name for z in uploaded_zips],
                            'Size (KB)': [f"{z.size / 1024:.1f}" for z in uploaded_zips]
                        }
                        pd.DataFrame(zip_data).to_excel(writer, index=False, sheet_name='ZIP Files')
                        
                        # Add skipped files sheet if any
                        if skipped:
                            pd.DataFrame(skipped, columns=['File', 'Error']).to_excel(writer, index=False, sheet_name='Errors')
                        
                        # Add no_keys files sheet if any
                        if no_keys:
                            pd.DataFrame({'File': no_keys}).to_excel(writer, index=False, sheet_name='No Keys')
                    
                    output.seek(0)
                    
                    st.download_button(
                        label=f"📥 Download Hasil (Excel) - {len(df_result)} rows",
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
            # Cleanup temp directories
            for temp_dir, _ in temp_dirs:
                if temp_dir and os.path.exists(temp_dir):
                    try:
                        shutil.rmtree(temp_dir)
                    except:
                        pass

else:
    st.info("👆 Silakan upload satu atau lebih file ZIP yang berisi folder dengan file Excel/ODS")
    
    # Show example
    with st.expander("💡 Panduan Lengkap"):
        st.markdown("""
        ### Langkah-langkah:
        
        1. **Persiapan Multiple Folder**
           - Folder 1: Rekap Januari
           - Folder 2: Rekap Februari
           - Folder 3: Rekap Maret
           - dst...
        
        2. **Compress Setiap Folder ke ZIP**
           - **Windows:** Klik kanan folder → Send to → Compressed (zipped) folder
           - **Mac:** Klik kanan folder → Compress
           - **Linux:** Klik kanan folder → Compress
           
           Hasilnya:
           - Rekap_Januari.zip
           - Rekap_Februari.zip
           - Rekap_Maret.zip
        
        3. **Upload Semua ZIP Sekaligus**
           - Click tombol **"Browse files"**
           - Select multiple files (Ctrl+Click atau Shift+Click)
           - Atau drag & drop semua ZIP files
        
        4. **Processing**
           - Click tombol **"Mulai Processing Semua ZIP"**
           - Tunggu hingga semua file dari semua ZIP diproses
        
        5. **Filter & Download**
           - Filter hasil berdasarkan ZIP tertentu (optional)
           - Download hasil lengkap dalam format Excel
        
        ### Contoh Struktur:
        ```
        📦 Rekap_Januari.zip
            └── Rekap Januari/
                ├── file1.xlsx
                └── file2.xls
        
        📦 Rekap_Februari.zip
            └── Rekap Februari/
                ├── file3.xlsx
                └── subfolder/
                    └── file4.ods
        
        📦 Rekap_Maret.zip
            └── Rekap Maret/
                └── file5.xlsm
        ```
        
        ### Output Excel akan berisi:
        - **Sheet "Data":** Hasil parse semua file dari semua ZIP
        - **Sheet "Summary":** Ringkasan statistik
        - **Sheet "ZIP Files":** Daftar ZIP yang diupload
        - **Sheet "Errors":** File yang error (jika ada)
        - **Sheet "No Keys":** File tanpa label (jika ada)
        
        ### Keuntungan Multiple ZIP:
        ✅ Upload banyak folder sekaligus  
        ✅ Source file otomatis diberi label nama ZIP  
        ✅ Bisa filter hasil berdasarkan ZIP tertentu  
        ✅ Lebih cepat daripada upload satu-satu  
        ✅ Ideal untuk rekap bulanan/tahunan  
        """)

# Footer
st.markdown("---")
st.markdown("Parser untuk E-Memo Rekap")
