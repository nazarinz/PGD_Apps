import io
import pandas as pd
import streamlit as st

def write_excel_autofit(df_map: dict[str, pd.DataFrame], date_cols: list[str] | None = None, file_name: str = "output.xlsx") -> bytes:
    """
    Buat file Excel dengan autofit columns, filter, dan freeze panes.
    
    Args:
        df_map: Dictionary dengan key=sheet_name, value=DataFrame
        date_cols: List kolom yang format sebagai tanggal (opsional)
        file_name: Nama file output
    
    Returns:
        bytes: Content Excel file yang siap di-download
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter", datetime_format="m/d/yyyy") as writer:
        for sheet, df in df_map.items():
            df.to_excel(writer, index=False, sheet_name=sheet)
            wb, ws = writer.book, writer.sheets[sheet]
            
            # Format untuk header
            header_fmt = wb.add_format({
                'bg_color': '#1f77b4',
                'font_color': 'white',
                'bold': True,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': True
            })
            
            # Format untuk data
            data_fmt = wb.add_format({
                'border': 1,
                'align': 'left',
                'valign': 'vcenter'
            })
            
            # Terapkan format header
            for col_num, value in enumerate(df.columns.values):
                ws.write(0, col_num, value, header_fmt)
            
            # Format untuk date columns
            if date_cols:
                date_fmt = wb.add_format({'num_format': 'm/d/yyyy', 'border': 1})
                for dc in date_cols:
                    if dc in df.columns:
                        idx = df.columns.get_loc(dc)
                        ws.set_column(idx, idx, 12, date_fmt)
            
            # Auto-fit columns dengan maksimal width
            for col_idx, col_name in enumerate(df.columns):
                sample_vals = df[col_name].head(1000).astype(str).tolist() if not df.empty else []
                maxlen = max([len(str(col_name))] + [len(s) for s in sample_vals])
                ws.set_column(col_idx, col_idx, min(50, max(10, maxlen + 2)), data_fmt)
            
            # Autofilter dan freeze panes
            ws.autofilter(0, 0, len(df), len(df.columns)-1)
            ws.freeze_panes(1, 0)
    
    return buf.getvalue()


def display_success_message(message: str):
    """Tampilkan pesan sukses dengan styling konsisten."""
    st.success(f"✅ {message}")


def display_error_message(message: str):
    """Tampilkan pesan error dengan styling konsisten."""
    st.error(f"❌ {message}")


def display_info_message(message: str):
    """Tampilkan pesan info dengan styling konsisten."""
    st.info(f"ℹ️ {message}")


def display_warning_message(message: str):
    """Tampilkan pesan warning dengan styling konsisten."""
    st.warning(f"⚠️ {message}")

