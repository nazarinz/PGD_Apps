from utils.auth import require_login

require_login()

# pages/3_Merger Daily Report.py
# Adapted from user's HTML-XLS merger app
import io
from datetime import datetime
from io import StringIO
from typing import List, Optional

import pandas as pd
import streamlit as st

st.set_page_config(page_title="PGD Apps — Merger Daily Report", page_icon="📚", layout="wide")
st.title("📚 Merger Daily Report")
st.caption("Dirancang untuk file *.xls* yang sebenarnya berisi HTML (sering dari export sistem).")

st.markdown('''
**Langkah:**
1) Upload banyak file `.xls` / `.html` (bisa banyak).
2) Klik **Proses**.
3) Lihat preview & unduh rekap Excel/CSV.

**Pembersihan yang dilakukan:**
- Ambil **tabel pertama** dari setiap file (`pandas.read_html`).
- Deteksi otomatis baris header, mendukung dua format:
  - Format baru: `fact_order, vbeln, auart, zzmdmark, zzmdnam, bstkd, sizeno, prod_qty`
  - Format lama: `FactOrder, Order no, Prod. Order Type, ...`
- Baris judul di atas header (mis. `ASB_Hourly_Productivity_Rep_Detail_...`) diabaikan.
- Drop baris kosong, header duplikat, dan baris yang kolom pertama mengandung **"Total"**.
- Ambil **8 kolom pertama** lalu ganti nama menjadi:
  `['FactOrder','Order no','Prod. Order Type','Article','Style Name','PO No','Size','Production Qty']`
- Tambah kolom `Source_File`.
''')

files = st.file_uploader(
    "Upload file .xls/.html (bisa banyak)",
    type=["xls", "html", "htm"],
    accept_multiple_files=True,
)
btn = st.button("🚀 Proses")

DEFAULT_COLS = ['FactOrder', 'Order no', 'Prod. Order Type', 'Article',
                'Style Name', 'PO No', 'Size', 'Production Qty']

# Nama header yang dikenali (format baru + format lama), dinormalisasi lowercase tanpa spasi/titik
HEADER_ALIASES = {
    "factorder", "fact_order",
    "orderno", "vbeln",
    "prodordertype", "auart",
    "article", "zzmdmark",
    "stylename", "zzmdnam",
    "pono", "bstkd",
    "size", "sizeno",
    "productionqty", "prod_qty",
}


def _norm(x) -> str:
    return (
        str(x).strip().lower()
        .replace(" ", "").replace(".", "")
        if pd.notna(x) else ""
    )


def read_html_tables_from_upload(f) -> List[pd.DataFrame]:
    if hasattr(f, "getvalue"):
        raw = f.getvalue()
    else:
        try:
            f.seek(0)
        except Exception:
            pass
        raw = f.read()

    for enc in ("utf-8", "latin-1", "cp1252"):
        try:
            text = raw.decode(enc, errors="ignore")
            # header=None -> semua baris (termasuk judul & header) dibaca sebagai data
            tables = pd.read_html(StringIO(text), header=None)
            if tables:
                return tables
        except Exception:
            continue
    try:
        tables = pd.read_html(io.BytesIO(raw), header=None)
        return tables
    except Exception:
        return []


def find_header_row(df: pd.DataFrame) -> Optional[int]:
    """Cari baris yang paling mirip header (>=4 sel cocok dengan alias)."""
    for i in range(min(len(df), 20)):
        hits = sum(_norm(v) in HEADER_ALIASES for v in df.iloc[i].tolist())
        if hits >= 4:
            return i
    return None


def is_header_like(row) -> bool:
    return sum(_norm(v) in HEADER_ALIASES for v in row) >= 4


def clean_text(s: pd.Series) -> pd.Series:
    s = s.astype("string").str.strip()
    s = s.replace({"": pd.NA, "nan": pd.NA, "NaN": pd.NA, "None": pd.NA, "<NA>": pd.NA})
    return s


def clean_id(s: pd.Series) -> pd.Series:
    """Bersihkan kolom ID/angka-teks (mis. 10198957.0 -> 10198957)."""
    s = clean_text(s)
    return s.str.replace(r"\.0+$", "", regex=True)


def process_table(raw_df: pd.DataFrame) -> pd.DataFrame:
    df = raw_df.copy()
    df = df.dropna(how="all").reset_index(drop=True)

    # Buang semua baris sampai header (judul laporan, dsb.)
    hdr_idx = find_header_row(df)
    if hdr_idx is not None:
        df = df.iloc[hdr_idx + 1:].reset_index(drop=True)

    # Buang header duplikat di tengah data
    df = df[~df.apply(lambda r: is_header_like(r.tolist()), axis=1)]

    # Pastikan minimal 8 kolom, ambil 8 kolom pertama
    if df.shape[1] < 8:
        for i in range(df.shape[1], 8):
            df[f"col_{i+1}"] = pd.NA
    df = df.iloc[:, :8].copy()
    df.columns = DEFAULT_COLS

    # Bersihkan nilai
    for c in ["FactOrder", "Prod. Order Type", "Article", "Style Name", "PO No", "Size"]:
        df[c] = clean_text(df[c])
    df["Order no"] = clean_id(df["Order no"])

    # Buang baris Total (kolom pertama mengandung "Total")
    df = df[~df["FactOrder"].fillna("").str.contains("total", case=False)]

    # Production Qty -> numerik
    df["Production Qty"] = pd.to_numeric(
        df["Production Qty"].astype("string").str.replace(",", "", regex=False).str.strip(),
        errors="coerce",
    )

    # Buang baris tanpa Order no & tanpa qty (sisa baris kosong/tidak valid)
    df = df[~(df["Order no"].isna() & df["Production Qty"].isna())]

    return df.reset_index(drop=True)


if btn:
    if not files:
        st.error("Upload minimal 1 file.")
        st.stop()

    frames = []
    log_rows = []
    for f in files:
        try:
            tables = read_html_tables_from_upload(f)
            if not tables:
                log_rows.append([f.name, "Gagal baca HTML", "-"])
                continue
            df = process_table(tables[0])
            df["Source_File"] = f.name
            frames.append(df)
            log_rows.append([f.name, "OK", f"{df.shape[0]} rows"])
        except Exception as e:
            log_rows.append([f.name, f"Error: {e}", "-"])

    if not frames:
        st.error("Tidak ada tabel yang berhasil dibaca dari file yang diupload.")
        st.dataframe(pd.DataFrame(log_rows, columns=["File", "Status", "Info"]), use_container_width=True)
        st.stop()

    combined = pd.concat(frames, ignore_index=True)
    st.success(f"Sukses gabung {len(frames)} file. Total baris: {combined.shape[0]}")

    st.subheader("🔎 Preview (Top 1000 rows)")
    st.dataframe(combined.head(1000), use_container_width=True)

    st.subheader("🧾 Log Baca File")
    log_df = pd.DataFrame(log_rows, columns=["File", "Status", "Info"])
    st.dataframe(log_df, use_container_width=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    buf_xlsx = io.BytesIO()
    with pd.ExcelWriter(buf_xlsx, engine="openpyxl") as writer:
        combined.to_excel(writer, index=False, sheet_name="Combined")
        log_df.to_excel(writer, index=False, sheet_name="Read_Log")
    buf_xlsx.seek(0)
    st.download_button(
        label="📥 Download Rekap (Excel)",
        data=buf_xlsx.getvalue(),
        file_name=f"rekap_html_xls_{ts}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    csv_data = combined.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Download Rekap (CSV)",
        data=csv_data,
        file_name=f"rekap_html_xls_{ts}.csv",
        mime="text/csv",
    )
