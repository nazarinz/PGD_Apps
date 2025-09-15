import io
import pandas as pd

def write_excel_autofit(df_map: dict[str, pd.DataFrame], date_cols: list[str] | None = None, file_name: str = "output.xlsx") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter", datetime_format="m/d/yyyy") as writer:
        for sheet, df in df_map.items():
            df.to_excel(writer, index=False, sheet_name=sheet)
            wb, ws = writer.book, writer.sheets[sheet]
            if date_cols:
                date_fmt = wb.add_format({'num_format': 'm/d/yyyy'})
                for dc in date_cols:
                    if dc in df.columns:
                        idx = df.columns.get_loc(dc)
                        ws.set_column(idx, idx, 12, date_fmt)
            for col_idx, col_name in enumerate(df.columns):
                sample_vals = df[col_name].head(1000).astype(str).tolist() if not df.empty else []
                maxlen = max([len(str(col_name))] + [len(s) for s in sample_vals])
                ws.set_column(col_idx, col_idx, min(50, max(10, maxlen + 2)))
            ws.autofilter(0, 0, len(df), len(df.columns)-1)
            ws.freeze_panes(1, 0)
    return buf.getvalue()

