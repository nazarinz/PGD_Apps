import streamlit as st
import pandas as pd
import io
import zipfile
import re
from datetime import datetime

st.set_page_config(page_title="Items relationship by PO", layout="wide")

REQUIRED_COLS = [
    "Order #",
    "Article Number",
    "Material Color Description",
    "Manufacturing Size",
    "UPC/EAN (GTIN)",
    "Country/Region"
]

COLUMN_RENAME = {
    "Article Number": "Style",
    "Material Color Description": "Color",
    "Manufacturing Size": "Size",
    "UPC/EAN (GTIN)": "Item No."
}

OUT_HEADERS = ["Style", "Color", "Size", "Item No."] + [f"A{i}" for i in range(1, 11)]

st.title("Items relationship table — STRICT column mode")

st.markdown(
    """
    Aplikasi ini **HANYA** menerima kolom berikut:

    - `Order #`
    - `Article Number`
    - `Material Color Description`
    - `Manufacturing Size`
    - `UPC/EAN (GTIN)`
    - `Country/Region`

    Semua kolom lain akan diabaikan.
    
    **Item No. yang 12 digit akan otomatis ditambahkan 0 di depan menjadi 13 digit.**
    """
)

uploaded_file = st.file_uploader("Upload file Excel (.xls/.xlsx)", type=["xls", "xlsx"])
sheet = None
if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheets = xls.sheet_names
        sheet = st.selectbox("Pilih sheet", sheets) if len(sheets) > 1 else sheets[0]
        df = pd.read_excel(xls, sheet_name=sheet, dtype=str)
    except Exception as e:
        st.error(f"Gagal membaca file Excel: {e}")
        st.stop()

    df.columns = df.columns.str.strip()
    df = df.fillna("").astype(str)
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # STRICT CHECK
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        st.error(f"Kolom berikut wajib ada (STRICT): {missing}")
        st.stop()

    # Keep only required columns (strict)
    df = df[REQUIRED_COLS].copy()

    # Rename to unified output format
    df = df.rename(columns=COLUMN_RENAME)

    # Auto-pad Item No. yang 12 digit menjadi 13 digit
    def pad_item_no(item):
        item = str(item).strip()
        # Hanya proses jika Item No. adalah angka dan panjangnya 12 digit
        if item.isdigit() and len(item) == 12:
            return "0" + item
        return item
    
    df["Item No."] = df["Item No."].apply(pad_item_no)

    st.subheader("Preview data (strict columns only)")
    st.dataframe(df.head(10))

    add_date_in_filename = st.checkbox("Tambahkan tanggal pada nama file", value=False)
    include_summary = st.checkbox("Masukkan summary.csv ke dalam ZIP", value=True)

    def safe_filename(s):
        s = re.sub(r'[\\/*?:"<>|]', "_", s)
        s = re.sub(r"\s+", " ", s).strip()
        return s

    def create_zip_bytes(df, add_date=False, include_summary=True):
        zip_buf = io.BytesIO()
        summary_rows = []

        with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            groups = df.groupby(["Order #", "Country/Region"], sort=False)

            for (order, country), group in groups:
                out_df = group[["Style", "Color", "Size", "Item No."]].copy()

                # Add A1–A10 blank columns
                for col in [f"A{i}" for i in range(1, 11)]:
                    out_df[col] = ""

                out_df = out_df.drop_duplicates().reset_index(drop=True)

                if out_df.empty:
                    continue

                name = f"Items relationship table_{order} {country}"
                if add_date:
                    name += f" {datetime.now().strftime('%Y%m%d')}"

                filename = safe_filename(name) + ".xlsx"

                excel_buf = io.BytesIO()
                out_df.to_excel(excel_buf, index=False, engine="openpyxl")
                excel_buf.seek(0)
                zf.writestr(filename, excel_buf.read())

                summary_rows.append({
                    "Order #": order,
                    "Country/Region": country,
                    "Filename": filename,
                    "Rows": len(out_df)
                })

            if include_summary:
                summary_df = pd.DataFrame(summary_rows)
                csv_bytes = summary_df.to_csv(index=False).encode("utf-8")
                zf.writestr("summary.csv", csv_bytes)

        zip_buf.seek(0)
        return zip_buf.getvalue(), summary_rows

    if st.button("Process & Generate ZIP"):
        try:
            with st.spinner("Memproses file..."):
                zip_bytes, summary_rows = create_zip_bytes(
                    df,
                    add_date=add_date_in_filename,
                    include_summary=include_summary
                )

            if len(summary_rows) == 0:
                st.warning("Tidak ada grup Order # + Country/Region ditemukan.")

            st.success("Selesai — klik untuk download ZIP")
            st.download_button(
                "Download ZIP",
                zip_bytes,
                file_name=f"items_relationship_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip"
            )

            if include_summary:
                st.subheader("Preview Summary")
                st.dataframe(pd.DataFrame(summary_rows))

        except Exception as e:
            st.error(f"Error: {e}")
