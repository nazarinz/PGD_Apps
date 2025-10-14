# app.py — PGD Subtotal Generator (FINAL v3 — robust WS=10 + hard color per-cell)
# Author: Nazarudin Zaini

import re
import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO

# ====================== Streamlit Config ======================
st.set_page_config(page_title="PGD Apps — Subtotal Generator", page_icon="📊", layout="wide")
st.title("📊 Subtotal Generator — Sizes, CRD_Mth, CRDPD_Mth")

# ====================== Upload Excel ======================
uploaded = st.file_uploader("📂 Upload file Excel", type=["xlsx", "xls"])
if uploaded is None:
    st.info("⬆️ Silakan upload file Excel terlebih dahulu")
    st.stop()

# Panduan untuk user baru
st.markdown("""
### 📑 Format File Excel (Minimal)
| Kolom | Keterangan |
|---|---|
| **Sales Order** | Nomor SO |
| **Document Date** | Tanggal dokumen order |
| **Article** | Hanya FG/HS dipakai (baris Article berawalan **HL** atau **HU** akan dihapus) |
| **Order Quantity** | Jumlah pesanan |
| **Working Status** | Status (gunakan 10 untuk new order) |
| **CRD**, **PD**, **LPD** | Tanggal penting |
| **UK_*** | Kolom ukuran (size breakdown) |
""")

df_sizelist = pd.read_excel(uploaded)

# Hapus baris dengan Article prefix HL/HU
if "Article" in df_sizelist.columns:
    df_sizelist = df_sizelist[~df_sizelist["Article"].astype(str).str.startswith(("HL", "HU"))]

# ====================== Input Parameter ======================
with st.expander("⚙️ Pengaturan Data"):
    new_order_date = st.date_input("📅 Tanggal New Order terakhir (WAJIB)")
    cancel_so_text = st.text_area("🗑️ Daftar Sales Order yang Cancel (pisahkan dengan koma atau enter):")
    cancel_sos = re.split(r"[,\s]+", cancel_so_text.strip()) if cancel_so_text else []
    cancel_sos = [s.strip() for s in cancel_sos if s.strip()]

execute_btn = st.button("🚀 Execute Process", type="primary")

if not execute_btn:
    st.warning("⚠️ Isi parameter lalu klik **Execute Process** untuk memproses data.")
    st.stop()

if not new_order_date:
    st.error("❗ Harap isi tanggal New Order terlebih dahulu.")
    st.stop()

NEW_ORDER_DATE = pd.to_datetime(new_order_date)

# ====================== Normalisasi tanggal ======================
for col in ["Document Date", "LPD", "CRD", "PD"]:
    if col in df_sizelist.columns:
        df_sizelist[col] = pd.to_datetime(df_sizelist[col], errors="coerce")

# ====================== Remark (robust WS=10) ======================
if "Remark" in df_sizelist.columns:
    df_sizelist.drop(columns=["Remark"], inplace=True)

# Coerce Working Status to numeric safely
ws_num = pd.to_numeric(df_sizelist.get("Working Status", pd.Series(index=df_sizelist.index)), errors="coerce")

remark = np.where(
    (df_sizelist["Document Date"] >= NEW_ORDER_DATE) &
    (df_sizelist["LPD"].isna()) &
    (ws_num == 10),   # robust check: 10 / 10.0 / "10"
    "New",
    "cfm"
)
insert_pos = df_sizelist.columns.get_loc("Document Date") + 1
df_sizelist.insert(insert_pos, "Remark", remark)

# ====================== Isi LPD kosong = min(CRD, PD) ======================
if {"CRD", "PD", "LPD"}.issubset(df_sizelist.columns):
    row_min = pd.concat([df_sizelist["CRD"], df_sizelist["PD"]], axis=1).min(axis=1, skipna=True)
    df_sizelist.loc[df_sizelist["LPD"].isna(), "LPD"] = row_min[df_sizelist["LPD"].isna()]

# ====================== Helpers ======================
def insert_after(df, after_col, new_col, values):
    pos = df.columns.get_loc(after_col) + 1
    df.insert(pos, new_col, values)

def bucket_base_from_day(day_series):
    return np.where(day_series >= 24, "30",
           np.where(day_series >= 16, "23",
           np.where(day_series >= 8,  "15",
           np.where(day_series >= 1,  "07", None))))

def fmt_mdy(dt):
    dt = pd.to_datetime(dt)
    return f"{dt.month}/{dt.day}/{dt.year}"

# ====================== CRD_Mth & CRDPD_Mth ======================
if "CRD" in df_sizelist.columns:
    YM_CRD   = df_sizelist["CRD"].dt.strftime("%Y%m")
    Day_CRD  = df_sizelist["CRD"].dt.day
    BucketB  = pd.Series(bucket_base_from_day(Day_CRD), index=df_sizelist.index, dtype="object")
    RemarkC  = df_sizelist["Remark"].str.lower().fillna("cfm")
    CRD_Mth  = YM_CRD.fillna("") + BucketB.fillna("") + "_" + RemarkC
    insert_after(df_sizelist, "CRD", "YM_CRD", YM_CRD)
    insert_after(df_sizelist, "YM_CRD", "Day_CRD", Day_CRD)
    insert_after(df_sizelist, "Day_CRD", "Class_CRD", BucketB)
    insert_after(df_sizelist, "Class_CRD", "CRD_Mth", CRD_Mth)

if "LPD" in df_sizelist.columns:
    YM_CRDPD  = df_sizelist["LPD"].dt.strftime("%Y%m")
    Day_CRDPD = df_sizelist["LPD"].dt.day
    BucketB2  = pd.Series(bucket_base_from_day(Day_CRDPD), index=df_sizelist.index, dtype="object")
    RemarkP   = df_sizelist["Remark"].str.lower().fillna("cfm")
    CRDPD_Mth = YM_CRDPD.fillna("") + BucketB2.fillna("") + "_" + RemarkP
    insert_after(df_sizelist, "LPD", "YM_CRDPD", YM_CRDPD)
    insert_after(df_sizelist, "YM_CRDPD", "Day_CRDPD", Day_CRDPD)
    insert_after(df_sizelist, "Day_CRDPD", "Class_CRDPD", BucketB2)
    insert_after(df_sizelist, "Class_CRDPD", "CRDPD_Mth", CRDPD_Mth)

# ====================== Subtotal builder (subtotal-only rows) ======================
size_cols = [c for c in df_sizelist.columns if re.match(r'(?i)^UK_', str(c))]
order_cols = ["Order Quantity"] + size_cols

def make_subtotal_only(df_source, group_col, order_cols, label_fmt):
    df_vis = df_source[[group_col] + order_cols].copy().sort_values(group_col, ascending=True, na_position="last")
    df_work = pd.concat(
        [df_source.loc[df_vis.index, ["Remark", "Sales Order"]].reset_index(drop=True),
         df_vis.reset_index(drop=True)],
        axis=1
    )

    pieces = []
    for key, grp in df_work.groupby(df_work[group_col], dropna=False, sort=True):
        subtotal = {col: "" for col in df_work.columns}
        label = "(blank)" if pd.isna(key) else label_fmt(key)
        # Subtotal is "New" only if there is at least one `New` row in that group
        has_new = (grp["Remark"].astype(str).str.lower() == "new").any()
        subtotal["Remark"] = "New" if has_new else "cfm"
        subtotal["Sales Order"] = ""  # not needed in subtotal rows
        subtotal[group_col] = label
        for col in order_cols:
            subtotal[col] = grp[col].sum(skipna=True)
        pieces.append(pd.DataFrame([subtotal], columns=df_work.columns))

    out = pd.concat(pieces, ignore_index=True)
    # Grand total (always cfm color)
    grand_vals = df_vis.drop(columns=[group_col]).sum(numeric_only=True)
    grand = {col: "" for col in out.columns}
    grand[group_col] = "Grand Total"
    grand["Remark"] = "cfm"
    for col in order_cols:
        grand[col] = grand_vals.get(col, "")
    out = pd.concat([out, pd.DataFrame([grand])], ignore_index=True)
    return out

sizes_df  = make_subtotal_only(df_sizelist, "Document Date", order_cols, label_fmt=lambda k: f"{fmt_mdy(k)} Total")
crd_df    = make_subtotal_only(df_sizelist, "CRD_Mth", order_cols, label_fmt=lambda k: str(k))
crdpd_df  = make_subtotal_only(df_sizelist, "CRDPD_Mth", order_cols, label_fmt=lambda k: str(k))
crdpd_df  = crdpd_df.rename(columns={"CRDPD_Mth": "CRDPD_month"})

# ====================== Preview ======================
st.success("✅ Data berhasil diproses!")
st.subheader("📑 Preview Data")
st.dataframe(df_sizelist.head(20), use_container_width=True)

# ====================== Excel export (hard color per-cell) ======================
def build_excel_bytes(df_sizelist, sizes_df, crd_df, crdpd_df, cancel_sos) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="m/d/yyyy") as writer:
        wb = writer.book

        # Base formats (no color)
        fmt_base  = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter"})
        fmt_date  = wb.add_format({"num_format":"m/d/yyyy","font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter"})
        fmt_header= wb.add_format({"bold":True,"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter"})
        fmt_num   = wb.add_format({"num_format":"0","font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter"})

        # Colored formats (preserve number formats!)
        fmt_base_red    = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","font_color":"#FF0000"})
        fmt_date_red    = wb.add_format({"num_format":"m/d/yyyy","font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","font_color":"#FF0000"})
        fmt_num_red     = wb.add_format({"num_format":"0","font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","font_color":"#FF0000"})
        fmt_base_purple = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","font_color":"#800080"})
        fmt_date_purple = wb.add_format({"num_format":"m/d/yyyy","font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","font_color":"#800080"})
        fmt_num_purple  = wb.add_format({"num_format":"0","font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","font_color":"#800080"})

        def autofit(ws, df):
            for j, c in enumerate(df.columns):
                maxlen = max(len(str(c)), df[c].astype(str).map(len).max() if len(df)>0 else 0)
                ws.set_column(j, j, max(8, min(60, maxlen + 2)))

        def cell_format_for(col_name, series_dtype, color=None):
            # Decide baseline
            is_dt = np.issubdtype(series_dtype, np.datetime64)
            is_num = (col_name == "Order Quantity") or bool(re.match(r'(?i)^UK_', str(col_name)))
            if color == "red":
                if is_dt:  return fmt_date_red
                if is_num: return fmt_num_red
                return fmt_base_red
            elif color == "purple":
                if is_dt:  return fmt_date_purple
                if is_num: return fmt_num_purple
                return fmt_base_purple
            else:
                if is_dt:  return fmt_date
                if is_num: return fmt_num
                return fmt_base

        def write_with_hard_color(ws, df, cancel_sos):
            """Rewrite colored rows per-cell to hard-code color without breaking date/num formats."""
            # First write via to_excel has already happened. Now re-write cells with formats.
            nrows, ncols = df.shape
            cols = list(df.columns)
            dtypes = [df[c].dtype for c in cols]
            # Find the index of 'Remark' and 'Sales Order' if exist (for color logic)
            has_remark = "Remark" in cols
            has_so = "Sales Order" in cols
            idx_remark = cols.index("Remark") if has_remark else None
            idx_so = cols.index("Sales Order") if has_so else None

            for i in range(nrows):
                # Determine row color:
                remark_val = str(df.iloc[i, idx_remark]).lower() if has_remark else ""
                so_val = str(df.iloc[i, idx_so]) if has_so else ""
                color = None
                if remark_val == "new":
                    color = "red"
                elif so_val in cancel_sos:
                    color = "purple"

                if color is None:
                    # leave as is (column formats already applied)
                    continue

                # Re-write every cell in the row with colored formats preserving number type
                for j in range(ncols):
                    val = df.iloc[i, j]
                    fmt = cell_format_for(cols[j], dtypes[j], color=color)

                    # Row 0 is header in Excel, data starts at row 1 offset:
                    excel_row = i + 1
                    excel_col = j

                    # Write respecting NaT/NaN
                    if pd.isna(val):
                        ws.write_blank(excel_row, excel_col, None, fmt)
                    else:
                        ws.write(excel_row, excel_col, val, fmt)

        # ---------- SHEET Data ----------
        df_sizelist.to_excel(writer, sheet_name="Data", index=False)
        ws = writer.sheets["Data"]
        nrow, ncol = df_sizelist.shape
        # Header first
        ws.set_row(0, None, fmt_header)
        # Apply column-level formats
        for j, c in enumerate(df_sizelist.columns):
            if np.issubdtype(df_sizelist[c].dtype, np.datetime64):
                ws.set_column(j, j, 12, fmt_date)
            elif c == "Order Quantity" or re.match(r'(?i)^UK_', str(c)):
                ws.set_column(j, j, 12, fmt_num)
            else:
                ws.set_column(j, j, 12, fmt_base)
        # Hard color after formats
        write_with_hard_color(ws, df_sizelist, cancel_sos)
        autofit(ws, df_sizelist)
        ws.freeze_panes(1, 0)

        # ---------- Subtotal sheets ----------
        for name, df_out in [("Sizes", sizes_df), ("CRD_Mth_Sizes", crd_df), ("CRDPD_Mth_Sizes", crdpd_df)]:
            df_out.to_excel(writer, sheet_name=name, index=False)
            wss = writer.sheets[name]
            wss.set_row(0, None, fmt_header)
            # Column formats (these sheets usually text + numbers)
            for j, c in enumerate(df_out.columns):
                if np.issubdtype(df_out[c].dtype, np.datetime64):
                    wss.set_column(j, j, 12, fmt_date)
                elif c == "Order Quantity" or re.match(r'(?i)^UK_', str(c)):
                    wss.set_column(j, j, 12, fmt_num)
                else:
                    wss.set_column(j, j, 12, fmt_base)
            # Hard color subtotal rows too (based on their Remark column)
            write_with_hard_color(wss, df_out, cancel_sos)
            autofit(wss, df_out)
            wss.freeze_panes(1, 0)

    return output.getvalue()

excel_bytes = build_excel_bytes(df_sizelist, sizes_df, crd_df, crdpd_df, cancel_sos)
st.download_button("⬇️ Download Excel (warna permanen)",
                   data=excel_bytes,
                   file_name="df_sizelist_ready.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("🔴 Merah = New Order (Document Date ≥ NEW_ORDER_DATE & LPD kosong & Working Status=10) • 🟣 Ungu = Sales Order termasuk daftar Cancel (indikasi visual, Remark tetap 'cfm')")
