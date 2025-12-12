# app.py
import streamlit as st
import pandas as pd
import io
import zipfile
import re
from datetime import datetime

st.set_page_config(page_title="Items relationship by PO", layout="wide")

st.title("Items relationship table — Split per Order # + Country/Region")
st.markdown(
    """
    Upload file Excel yang berisi kolom minimal `Order #` dan `Country/Region`.
    Aplikasi akan menghasilkan file Excel per kombinasi Order # + Country/Region dengan nama:
    `Items relationship table_<Order#> <Country/Region>.xlsx` dan mengemas semuanya ke ZIP.
    """
)

# instructions / help
with st.expander("Contoh kolom input yang dikenali (bisa variasi)"):
    st.write("""
    - Order # (wajib)
    - Country/Region (wajib)
    - Article Number / Style / Article / Model No.  -> dipetakan ke `Style`
    - Material Color Description / Color -> dipetakan ke `Color`
    - Manufacturing Size / Size -> dipetakan ke `Size`
    - UPC/EAN (GTIN) / UPC/EAN / GTIN / Item No -> dipetakan ke `Item No.`
    """)

uploaded_file = st.file_uploader("Upload file Excel (.xls/.xlsx)", type=["xls", "xlsx"])
sheet = None
if uploaded_file is not None:
    try:
        # try to read all sheets list for selection
        xls = pd.ExcelFile(uploaded_file)
        sheets = xls.sheet_names
        if len(sheets) > 1:
            sheet = st.selectbox("Pilih sheet", sheets)
        else:
            sheet = sheets[0]
        # read selected sheet
        df = pd.read_excel(xls, sheet_name=sheet, dtype=str)
    except Exception as e:
        st.error(f"Gagal membaca file Excel: {e}")
        st.stop()

    # normalize
    df.columns = df.columns.str.strip()
    df = df.fillna("").astype(str)
    df = df.applymap(lambda s: s.strip() if isinstance(s, str) else s)

    st.subheader("Preview data (5 baris pertama)")
    st.dataframe(df.head(5))

    # check grouping columns
    if "Order #" not in df.columns or "Country/Region" not in df.columns:
        st.error("File harus mengandung kolom `Order #` dan `Country/Region`.")
        st.stop()

    # options
    add_date_in_filename = st.checkbox("Tambahkan tanggal pada nama file (YYYYMMDD)", value=False)
    include_summary = st.checkbox("Masukkan file summary (summary.csv) di dalam zip", value=True)

    st.markdown("---")
    st.write("Mapping kolom (otomatis, akan disesuaikan jika nama kolom berbeda):")

    # mapping candidates
    col_map_candidates = {
        "Style": ["Article Number", "Style", "Article", "Model No", "Model No."],
        "Color": ["Material Color Description", "Color", "Material Color"],
        "Size": ["Manufacturing Size", "Size", "Manufacturing Size/Width"],
        "Item No.": ["UPC/EAN (GTIN)", "UPC/EAN", "GTIN", "Item No", "Item No."]
    }

    # choose mapping and show to user
    col_map = {}
    for out_col, cand_list in col_map_candidates.items():
        chosen = None
        for cand in cand_list:
            if cand in df.columns:
                chosen = cand
                break
        col_map[out_col] = chosen
        st.write(f"- {out_col}  ⟵  `{chosen or '--- (kosong / akan diisi blank)'}`")

    # output headers
    out_headers = ["Style", "Color", "Size", "Item No."] + [f"A{i}" for i in range(1, 11)]

    def safe_filename(s: str) -> str:
        # replace illegal filename chars
        s = re.sub(r'[\\/*?:"<>|]', '_', s)
        s = re.sub(r'\s+', ' ', s).strip()
        return s

    def create_zip_bytes(df: pd.DataFrame, add_date=False, include_summary=True):
        # create in-memory zip
        zip_buf = io.BytesIO()
        summary_rows = []
        with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            groups = df.groupby(["Order #", "Country/Region"], sort=False)
            for (order, country), group in groups:
                data = {}
                # for each output column, map from source or fill blanks
                for out_col in out_headers:
                    src = col_map.get(out_col)
                    if src:
                        # ensure lengths equal to group
                        data[out_col] = group[src].astype(str).tolist()
                    else:
                        data[out_col] = [""] * len(group)
                out_df = pd.DataFrame(data).drop_duplicates().reset_index(drop=True)

                if out_df.empty:
                    continue

                # filename
                name = f"Items relationship table_{order} {country}"
                if add_date:
                    name = f"{name} {datetime.now().strftime('%Y%m%d')}"
                filename = safe_filename(name) + ".xlsx"

                # write excel into buffer
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

            # add summary if requested
            if include_summary:
                summary_df = pd.DataFrame(summary_rows)
                if summary_df.empty:
                    summary_df = pd.DataFrame(columns=["Order #", "Country/Region", "Filename", "Rows"])
                csv_buf = summary_df.to_csv(index=False).encode("utf-8")
                zf.writestr("summary.csv", csv_buf)

        zip_buf.seek(0)
        return zip_buf.getvalue(), summary_rows

    if st.button("Process & Generate ZIP"):
        try:
            with st.spinner("Membuat file..."):
                zip_bytes, summary_rows = create_zip_bytes(df, add_date=add_date_in_filename, include_summary=include_summary)
            if len(summary_rows) == 0:
                st.warning("Tidak ditemukan grup Order # + Country/Region atau semua output kosong.")
            st.success("Selesai — ZIP siap di-download.")
            st.download_button(
                label="Download ZIP file",
                data=zip_bytes,
                file_name=f"items_relationship_by_po_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip"
            )

            if include_summary and summary_rows:
                st.subheader("Summary (preview)")
                st.dataframe(pd.DataFrame(summary_rows).head(50))
        except Exception as e:
            st.error(f"Terjadi error saat proses: {e}")
