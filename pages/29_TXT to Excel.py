import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="TXT Converter", layout="wide")

st.title("📄 TXT to Excel / CSV Converter")
st.markdown("Upload TXT SAP lalu download menjadi Excel, CSV, atau XLSB.")

# =========================================
# Upload File
# =========================================
uploaded_file = st.file_uploader("Upload TXT File", type=["txt"])

if uploaded_file is not None:

    try:

        # =========================================
        # AUTO DETECT ENCODING
        # =========================================
        encodings = ["utf-8", "latin1", "cp1252", "ISO-8859-1"]

        df = None

        for enc in encodings:

            try:
                uploaded_file.seek(0)

                df = pd.read_csv(
                    uploaded_file, sep="\t", encoding=enc, dtype=str, low_memory=False
                )

                st.success(f"✅ File berhasil dibaca menggunakan encoding: {enc}")
                break

            except:
                continue

        if df is None:
            st.error("❌ Gagal membaca file")
            st.stop()

        # =========================================
        # REMOVE UNNAMED COLUMN
        # =========================================
        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

        # =========================================
        # INFO
        # =========================================
        col1, col2 = st.columns(2)

        with col1:
            st.metric("Rows", len(df))

        with col2:
            st.metric("Columns", len(df.columns))

        # =========================================
        # PREVIEW
        # =========================================
        st.subheader("📊 Preview Data")

        st.dataframe(df, use_container_width=True, height=600)

        st.divider()

        st.subheader("⬇ Download File")

        # =========================================
        # DOWNLOAD EXCEL XLSX
        # =========================================
        excel_buffer = BytesIO()

        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:

            df.to_excel(writer, index=False, sheet_name="SAP_Data")

            worksheet = writer.sheets["SAP_Data"]

            # AUTO WIDTH
            for column_cells in worksheet.columns:

                max_length = 0
                column = column_cells[0].column_letter

                for cell in column_cells:
                    try:
                        max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass

                worksheet.column_dimensions[column].width = min(max_length + 5, 50)

        excel_buffer.seek(0)

        st.download_button(
            label="⬇ Download Excel (.xlsx)",
            data=excel_buffer,
            file_name="SAP_Converted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # =========================================
        # DOWNLOAD CSV
        # =========================================
        csv = df.to_csv(index=False).encode("utf-8")

        st.download_button(
            label="⬇ Download CSV (.csv)",
            data=csv,
            file_name="SAP_Converted.csv",
            mime="text/csv",
        )

        # =========================================
        # DOWNLOAD XLSB
        # =========================================
        st.info("⚠ XLSB export membutuhkan library tambahan: pyxlsb dan pyxlsbwriter")

        try:
            import pyxlsbwriter

            xlsb_buffer = BytesIO()

            with pyxlsbwriter.Workbook(xlsb_buffer) as workbook:

                sheet = workbook.add_worksheet("SAP_Data")

                # HEADER
                for col_num, value in enumerate(df.columns):
                    sheet.write(0, col_num, value)

                # DATA
                for row_num, row in enumerate(df.values, start=1):
                    for col_num, value in enumerate(row):
                        sheet.write(row_num, col_num, str(value))

            xlsb_buffer.seek(0)

            st.download_button(
                label="⬇ Download XLSB (.xlsb)",
                data=xlsb_buffer,
                file_name="SAP_Converted.xlsb",
                mime="application/vnd.ms-excel.sheet.binary.macroEnabled.12",
            )

        except:
            st.warning(
                "Install library XLSB terlebih dahulu:\n\npip install pyxlsbwriter"
            )

    except Exception as e:
        st.error(f"❌ Error: {e}")
