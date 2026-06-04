from utils.auth import require_login

require_login()

import streamlit as st
import os
import glob
import traceback
import shutil
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Tuple
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

EXCEL_EXTENSIONS = {".xls", ".xlsx", ".xlsm", ".xlsb", ".ods"}

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

def parse_excel_file(path: str) -> Dict[str, str]:
    df = safe_read_any(path)
    collected = {}
    for i in range(min(len(df), 400)):
        try:
            m = row_to_mapping(df.iloc[i, :].tolist())
        except Exception:
            m = {}
        for k, v in m.items():
            if k not in collected and v is not None:
                collected[k] = v
        if set(collected.keys()) >= KEY_SET:
            break
    return {col: collected.get(col, "") for col in TARGET_HEADERS}

def extract_zip_to_temp(zip_file) -> str:
    """Extract uploaded ZIP to a new temporary directory, return temp dir path."""
    temp_dir = tempfile.mkdtemp()
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)
    return temp_dir

def find_excel_files_in_dir(root_dir: str) -> List[str]:
    """Find all Excel/ODS files recursively inside a directory."""
    files = []
    for ext in ["*.xls", "*.xlsx", "*.xlsm", "*.xlsb", "*.ods"]:
        files += glob.glob(os.path.join(root_dir, "**", ext), recursive=True)
    return sorted(files)

def save_uploaded_excel_to_temp(uploaded_file) -> str:
    """Save a single uploaded Excel file to a temp file and return its path."""
    suffix = Path(uploaded_file.name).suffix
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.read())
    tmp.close()
    return tmp.name

def classify_uploads(uploaded_files) -> Tuple[List, List]:
    """Split uploaded files into ZIP list and Excel list."""
    zips = []
    excels = []
    for f in uploaded_files:
        ext = Path(f.name).suffix.lower()
        if ext == ".zip":
            zips.append(f)
        elif ext in EXCEL_EXTENSIONS:
            excels.append(f)
    return zips, excels

def process_all_uploads(uploaded_files):
    """
    Process a mixed list of ZIP and Excel uploads.
    Returns (df_result, no_keys, skipped, file_list, temp_dirs_to_cleanup)
    """
    zips, excels = classify_uploads(uploaded_files)

    all_rows = []
    all_skipped = []
    all_no_keys = []
    all_file_list = []   # display labels
    temp_dirs = []       # for cleanup

    today_date = pd.Timestamp.today().normalize()

    # ── 1. Extract ZIP files ────────────────────────────────────────────────
    if zips:
        st.info(f"📦 Mengekstrak {len(zips)} file ZIP...")
        zip_progress = st.progress(0)
        extracted = []
        for idx, zf in enumerate(zips):
            try:
                td = extract_zip_to_temp(zf)
                temp_dirs.append(td)
                extracted.append((td, zf.name))
                zip_progress.progress((idx + 1) / len(zips))
            except Exception as e:
                st.error(f"❌ Gagal ekstrak {zf.name}: {e}")
        st.success(f"✅ {len(extracted)} ZIP berhasil diekstrak")
    else:
        extracted = []

    # ── 2. Save direct Excel uploads to temp files ──────────────────────────
    excel_temps = []   # list of (temp_path, original_name)
    if excels:
        st.info(f"📄 Menyiapkan {len(excels)} file Excel/ODS...")
        for ef in excels:
            try:
                tp = save_uploaded_excel_to_temp(ef)
                excel_temps.append((tp, ef.name))
            except Exception as e:
                st.error(f"❌ Gagal menyimpan {ef.name}: {e}")

    # ── 3. Collect all (file_path, display_label) pairs ────────────────────
    work_items: List[Tuple[str, str]] = []

    for td, zip_name in extracted:
        for fp in find_excel_files_in_dir(td):
            rel = os.path.relpath(fp, td)
            label = f"[{zip_name}] {rel}"
            work_items.append((fp, label))

    for tp, orig_name in excel_temps:
        work_items.append((tp, orig_name))

    all_file_list = [label for _, label in work_items]

    if not work_items:
        return pd.DataFrame(columns=OUTPUT_COLS), [], [], [], temp_dirs, excel_temps

    st.info(f"🔍 Total {len(work_items)} file Excel/ODS akan diproses")

    # ── 4. Parse each file ──────────────────────────────────────────────────
    progress_bar = st.progress(0)
    status_text = st.empty()

    for idx, (fp, label) in enumerate(work_items):
        status_text.text(f"({idx+1}/{len(work_items)}) {label}")
        progress_bar.progress((idx + 1) / len(work_items))
        try:
            row = parse_excel_file(fp)
            all_empty = all(not str(row[c]).strip() for c in TARGET_HEADERS)
            if all_empty:
                all_no_keys.append(label)
                continue
            row["_source_file"] = label
            row["Tanggal Rekap"] = today_date
            all_rows.append(row)
        except Exception as e:
            all_skipped.append((label, str(e)))

    status_text.text("✅ Selesai!")

    df_out = (
        pd.DataFrame(all_rows).reindex(columns=OUTPUT_COLS)
        if all_rows
        else pd.DataFrame(columns=OUTPUT_COLS)
    )

    return df_out, all_no_keys, all_skipped, all_file_list, temp_dirs, excel_temps


# ────────────────────────────────────────────────────────────────────────────
# Streamlit UI
# ────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Excel/ODS Parser", page_icon="📊", layout="wide")

st.title("📊 Excel/ODS Parser untuk E-Memo")
st.markdown(
    "Upload **file Excel/ODS** secara langsung **atau** **file ZIP** yang berisi folder "
    "dengan file Excel/ODS — bisa dicampur sekaligus."
)

# ── Sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("ℹ️ Informasi")
    st.markdown("""
**Format yang didukung:**
- .xls, .xlsx, .xlsm, .xlsb, .ods *(upload langsung)*
- .zip *(berisi file Excel/ODS)*

**Bisa dicampur:** upload ZIP dan Excel sekaligus ✅

**Target Headers:**
""")
    with st.expander("Lihat semua headers"):
        for h in TARGET_HEADERS:
            st.text(f"• {h}")

# ── Upload area ──────────────────────────────────────────────────────────────
st.info(
    "💡 **Cara Upload:** Pilih satu atau lebih file. "
    "Boleh campur file Excel (.xlsx, .xls, .ods, dll.) dan file ZIP dalam satu kali upload."
)

uploaded_files = st.file_uploader(
    "Upload File Excel / ZIP",
    type=["zip", "xls", "xlsx", "xlsm", "xlsb", "ods"],
    accept_multiple_files=True,
    help="Pilih sembarang kombinasi file Excel/ODS dan/atau ZIP yang berisi file Excel/ODS.",
)

if uploaded_files:
    zips_up, excels_up = classify_uploads(uploaded_files)

    # Summary badge
    badges = []
    if zips_up:
        badges.append(f"📦 {len(zips_up)} ZIP")
    if excels_up:
        badges.append(f"📄 {len(excels_up)} Excel/ODS")
    st.success(f"✅ {len(uploaded_files)} file dipilih — {' · '.join(badges)}")

    with st.expander("📋 Daftar file yang dipilih"):
        for f in uploaded_files:
            ext = Path(f.name).suffix.lower()
            icon = "📦" if ext == ".zip" else "📄"
            st.text(f"{icon} {f.name}  ({f.size / 1024:.1f} KB)")

    if st.button("🚀 Mulai Processing", type="primary"):
        temp_dirs = []
        excel_temps = []
        try:
            df_result, no_keys, skipped, file_list, temp_dirs, excel_temps = process_all_uploads(uploaded_files)

            # ── Metrics ──────────────────────────────────────────────────────
            st.header("📈 Hasil Rekap")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("File ZIP", len(zips_up))
            c2.metric("File Excel/ODS Langsung", len(excels_up))
            c3.metric("Berhasil Diparse", len(df_result))
            c4.metric("Error + Tanpa Label", len(skipped) + len(no_keys))

            # ── File list ─────────────────────────────────────────────────────
            with st.expander(f"📁 Semua file yang diproses ({len(file_list)})"):
                for f in file_list:
                    st.text(f"  • {f}")

            # ── Data table ────────────────────────────────────────────────────
            if not df_result.empty:
                st.subheader("📋 Data Hasil Parse")

                # Filter by source ZIP (only show if there are ZIPs)
                if zips_up:
                    filter_options = ["Semua"] + [f"[{z.name}]" for z in zips_up] + (["[Excel Langsung]"] if excels_up else [])
                    selected_filter = st.selectbox("Filter sumber:", filter_options)

                    if selected_filter == "Semua":
                        filtered_df = df_result
                    elif selected_filter == "[Excel Langsung]":
                        filtered_df = df_result[~df_result["_source_file"].str.startswith("[")]
                    else:
                        filtered_df = df_result[df_result["_source_file"].str.startswith(selected_filter)]
                else:
                    filtered_df = df_result

                st.dataframe(filtered_df, use_container_width=True, height=420)
                st.caption(f"Menampilkan {len(filtered_df)} dari {len(df_result)} baris data")

                # ── Download ──────────────────────────────────────────────────
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_filename = f"Result_{timestamp}.xlsx"

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_result.to_excel(writer, index=False, sheet_name="Data")

                    # Summary sheet
                    pd.DataFrame({
                        "Metric": ["Total File ZIP", "Total File Excel Langsung", "Total File Ditemukan",
                                   "Berhasil Diparse", "Tanpa Label", "Error"],
                        "Count": [len(zips_up), len(excels_up), len(file_list),
                                  len(df_result), len(no_keys), len(skipped)],
                    }).to_excel(writer, index=False, sheet_name="Summary")

                    # Uploaded files list
                    pd.DataFrame({
                        "File": [f.name for f in uploaded_files],
                        "Tipe": ["ZIP" if Path(f.name).suffix.lower() == ".zip" else "Excel/ODS" for f in uploaded_files],
                        "Ukuran (KB)": [f"{f.size / 1024:.1f}" for f in uploaded_files],
                    }).to_excel(writer, index=False, sheet_name="File Diupload")

                    if skipped:
                        pd.DataFrame(skipped, columns=["File", "Error"]).to_excel(
                            writer, index=False, sheet_name="Errors"
                        )
                    if no_keys:
                        pd.DataFrame({"File": no_keys}).to_excel(
                            writer, index=False, sheet_name="No Keys"
                        )

                output.seek(0)
                st.download_button(
                    label=f"📥 Download Hasil Excel — {len(df_result)} baris",
                    data=output,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            else:
                st.warning("⚠️ Tidak ada data yang berhasil diparse dari file yang diupload.")

            # ── No-keys files ─────────────────────────────────────────────────
            if no_keys:
                with st.expander(f"⚠️ File tanpa label yang cocok ({len(no_keys)})"):
                    for f in no_keys:
                        st.text(f"• {f}")

            # ── Error files ───────────────────────────────────────────────────
            if skipped:
                with st.expander(f"❌ File dengan error ({len(skipped)})"):
                    for fname, err in skipped[:10]:
                        st.text(f"• {fname}")
                        st.code(err, language=None)
                    if len(skipped) > 10:
                        st.text(f"... dan {len(skipped) - 10} error lainnya")

        except Exception as e:
            st.error(f"❌ Terjadi error tak terduga: {e}")
            st.code(traceback.format_exc())

        finally:
            # Cleanup temp directories from ZIPs
            for td in temp_dirs:
                if td and os.path.exists(td):
                    try:
                        shutil.rmtree(td)
                    except Exception:
                        pass
            # Cleanup temp files from direct Excel uploads
            for tp, _ in excel_temps:
                if tp and os.path.exists(tp):
                    try:
                        os.remove(tp)
                    except Exception:
                        pass

else:
    st.info("👆 Upload satu atau lebih file Excel/ODS dan/atau ZIP di atas untuk memulai.")

    with st.expander("💡 Panduan Lengkap"):
        st.markdown("""
### Mode Upload yang Didukung

| Mode | Caranya | Cocok untuk |
|------|---------|-------------|
| **Excel Langsung** | Upload .xlsx / .xls / .ods dll. | File satuan, tidak perlu zip dulu |
| **ZIP** | Upload .zip berisi folder/file Excel | Banyak file dalam satu folder |
| **Campuran** | Upload ZIP + Excel bersamaan | Fleksibel, sesuai kebutuhan |

### Cara Upload Campuran:
1. Klik **"Browse files"**
2. Pilih kombinasi file ZIP dan/atau Excel (Ctrl+Click / Shift+Click)
3. Klik **"Mulai Processing"**

### Contoh Struktur ZIP:
```
📦 Rekap_Januari.zip
    └── Rekap Januari/
        ├── file1.xlsx
        └── subfolder/
            └── file2.ods

📄 file_tambahan.xlsx   ← upload langsung, tanpa perlu zip
```

### Output Excel berisi:
- **Data** — hasil parse semua file
- **Summary** — ringkasan statistik
- **File Diupload** — daftar file yang diupload (ZIP & Excel)
- **Errors** — file yang error *(jika ada)*
- **No Keys** — file tanpa label yang cocok *(jika ada)*
""")

st.markdown("---")
st.markdown("Parser untuk E-Memo Rekap")
