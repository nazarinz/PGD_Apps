# pages/3_Merger Daily Report.py
# Adapted from user's HTML-XLS merger app
import io
from datetime import datetime
from typing import List
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
- Drop baris kosong & baris yang kolom pertama = `FactOrder` (header duplikat).
- Drop baris yang kolom pertama mengandung kata **"Total"**.
- Ambil **8 kolom pertama**.
- Set nama kolom menjadi: `['FactOrder','Order no','Prod. Order Type','Article','Style Name','PO No','Size','Production Qty']`
- Tambah kolom `Source_File`.
''')

files = st.file_uploader("Upload file .xls/.html (bisa banyak)", type=["xls","html","htm"], accept_multiple_files=True)
btn = st.button("🚀 Proses")

DEFAULT_COLS = ['FactOrder', 'Order no', 'Prod. Order Type', 'Article',
                'Style Name', 'PO No', 'Size', 'Production Qty']

def read_html_tables_from_upload(f) -> List[pd.DataFrame]:
    raw = None
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
            tables = pd.read_html(text)
            if tables and len(tables) > 0:
                return tables
        except Exception:
            continue
    try:
        tables = pd.read_html(io.BytesIO(raw))
        return tables
    except Exception:
        return []

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
            df = tables[0].copy()
            df = df.dropna(how='all')
            first_col = df.columns[0]
            df = df[df[first_col].astype(str) != "FactOrder"]
            df = df[~df[first_col].astype(str).str.contains("Total", na=False, case=False)]
            if df.shape[1] < 8:
                for i in range(df.shape[1], 8):
                    df[f"col_{i+1}"] = pd.NA
            df = df.iloc[:, :8]
            df.columns = DEFAULT_COLS
            df["Source_File"] = f.name
            frames.append(df)
            log_rows.append([f.name, "OK", f"{df.shape[0]} rows"])
        except Exception as e:
            log_rows.append([f.name, f"Error: {e}", "-"])

    if not frames:
        st.error("Tidak ada tabel yang berhasil dibaca dari file yang diupload.")
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