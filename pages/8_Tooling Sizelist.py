# app.py — PGD Subtotal Generator (FINAL v2 — Execute Button Mode)
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

# Panduan
st.markdown("""
### 📑 Format File Excel
Pastikan file memiliki kolom berikut:

| Kolom | Keterangan |
|-------|-------------|
| **Sales Order** | Nomor SO |
| **Document Date** | Tanggal dokumen order |
| **Article** | Hanya FG/HS (HL, HU dihapus) |
| **Order Quantity** | Jumlah pesanan |
| **Working Status** | Status produksi (gunakan '10' untuk new order) |
| **CRD**, **PD**, **LPD** | Tanggal penting |
| **UK_*** | Kolom ukuran (size breakdown) |
""")

df_sizelist = pd.read_excel(uploaded)

# Filter hanya FG/HS
if "Article" in df_sizelist.columns:
    df_sizelist = df_sizelist[~df_sizelist["Article"].astype(str).str.startswith(("HL", "HU"))]

# ====================== Input Parameter ======================
with st.expander("⚙️ Pengaturan Data"):
    new_order_date = st.date_input("📅 Tanggal New Order terakhir (WAJIB)")
    cancel_so_text = st.text_area("🗑️ Masukkan daftar Sales Order yang Cancel (pisahkan dengan koma atau enter):")
    cancel_sos = re.split(r"[,\s]+", cancel_so_text.strip()) if cancel_so_text else []
    cancel_sos = [s.strip() for s in cancel_sos if s.strip()]

execute_btn = st.button("🚀 Execute Process", type="primary")

# Stop sampai tombol diklik
if not execute_btn:
    st.warning("⚠️ Silakan isi parameter lalu klik **Execute Process** untuk memproses data.")
    st.stop()

# Validasi input
if not new_order_date:
    st.error("❗ Harap isi tanggal New Order terlebih dahulu.")
    st.stop()

NEW_ORDER_DATE = pd.to_datetime(new_order_date)

# ====================== Normalisasi tanggal ======================
for col in ["Document Date", "LPD", "CRD", "PD"]:
    if col in df_sizelist.columns:
        df_sizelist[col] = pd.to_datetime(df_sizelist[col], errors="coerce")

# ====================== Remark Logic ======================
if "Remark" in df_sizelist.columns:
    df_sizelist.drop(columns=["Remark"], inplace=True)

remark = np.select(
    [
        (df_sizelist["Document Date"] >= NEW_ORDER_DATE)
        & (df_sizelist["LPD"].isna())
        & (df_sizelist["Working Status"].astype(str).str.strip() == "10"),
    ],
    ["New"],
    default="cfm"
)
insert_pos = df_sizelist.columns.get_loc("Document Date") + 1
df_sizelist.insert(insert_pos, "Remark", remark)

# ====================== Isi LPD kosong ======================
if {"CRD","PD","LPD"}.issubset(df_sizelist.columns):
    row_min = pd.concat([df_sizelist["CRD"], df_sizelist["PD"]], axis=1).min(axis=1, skipna=True)
    df_sizelist.loc[df_sizelist["LPD"].isna(), "LPD"] = row_min[df_sizelist["LPD"].isna()]

# ====================== Helper ======================
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
    Remark   = df_sizelist["Remark"].str.lower().fillna("cfm")
    CRD_Mth  = YM_CRD.fillna("") + BucketB.fillna("") + "_" + Remark
    insert_after(df_sizelist, "CRD", "YM_CRD", YM_CRD)
    insert_after(df_sizelist, "YM_CRD", "Day_CRD", Day_CRD)
    insert_after(df_sizelist, "Day_CRD", "Class_CRD", BucketB)
    insert_after(df_sizelist, "Class_CRD", "CRD_Mth", CRD_Mth)

if "LPD" in df_sizelist.columns:
    YM_CRDPD  = df_sizelist["LPD"].dt.strftime("%Y%m")
    Day_CRDPD = df_sizelist["LPD"].dt.day
    BucketB2  = pd.Series(bucket_base_from_day(Day_CRDPD), index=df_sizelist.index, dtype="object")
    Remark2   = df_sizelist["Remark"].str.lower().fillna("cfm")
    CRDPD_Mth = YM_CRDPD.fillna("") + BucketB2.fillna("") + "_" + Remark2
    insert_after(df_sizelist, "LPD", "YM_CRDPD", YM_CRDPD)
    insert_after(df_sizelist, "YM_CRDPD", "Day_CRDPD", Day_CRDPD)
    insert_after(df_sizelist, "Day_CRDPD", "Class_CRDPD", BucketB2)
    insert_after(df_sizelist, "Class_CRDPD", "CRDPD_Mth", CRDPD_Mth)

# ====================== Subtotal builder ======================
size_cols = [c for c in df_sizelist.columns if re.match(r'(?i)^UK_', str(c))]
order_cols = ["Order Quantity"] + size_cols

def make_subtotal_only(df_source, group_col, order_cols, label_fmt):
    df_vis = df_source[[group_col] + order_cols].copy().sort_values(group_col, ascending=True, na_position="last")
    df_work = pd.concat([df_source.loc[df_vis.index, ["Remark"]].reset_index(drop=True),
                         df_vis.reset_index(drop=True)], axis=1)
    pieces = []
    for key, grp in df_work.groupby(df_work[group_col], dropna=False, sort=True):
        subtotal = {col:"" for col in df_work.columns}
        label = "(blank)" if pd.isna(key) else label_fmt(key)
        has_new = (grp["Remark"].astype(str).str.lower()=="new").any()
        subtotal["Remark"] = "New" if has_new else "cfm"
        subtotal[group_col] = label
        for col in order_cols: subtotal[col] = grp[col].sum(skipna=True)
        pieces.append(pd.DataFrame([subtotal], columns=df_work.columns))
    out = pd.concat(pieces, ignore_index=True)
    grand_vals = df_vis.drop(columns=[group_col]).sum(numeric_only=True)
    grand = {col:"" for col in out.columns}
    grand[group_col] = "Grand Total"; grand["Remark"] = "cfm"
    for col in order_cols: grand[col] = grand_vals.get(col, "")
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

# ====================== Export Excel ======================
def build_excel_bytes(df_sizelist, sizes_df, crd_df, crdpd_df, cancel_sos) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="m/d/yyyy") as writer:
        wb = writer.book
        fmt_base  = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter"})
        fmt_date  = wb.add_format({"num_format":"m/d/yyyy","font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter"})
        fmt_header= wb.add_format({"bold":True,"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter"})
        fmt_red   = wb.add_format({"font_color":"#FF0000","font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter"})
        fmt_purple= wb.add_format({"font_color":"#800080","font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter"})
        fmt_num   = wb.add_format({"num_format":"0","font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter"})

        def autofit(ws, df):
            for j, c in enumerate(df.columns):
                maxlen = max(len(str(c)), df[c].astype(str).map(len).max() if len(df)>0 else 0)
                ws.set_column(j, j, max(8, min(60, maxlen + 2)))

        def color_rows(ws, df):
            for i in range(len(df)):
                remark_val = str(df.iloc[i].get("Remark", "")).lower()
                so_val = str(df.iloc[i].get("Sales Order", ""))
                fmt_row = fmt_base
                if remark_val == "new":
                    fmt_row = fmt_red
                elif so_val in cancel_sos:
                    fmt_row = fmt_purple
                ws.set_row(i + 1, None, fmt_row)

        # Sheet utama
        df_sizelist.to_excel(writer, sheet_name="Data", index=False)
        ws = writer.sheets["Data"]
        nrow, ncol = df_sizelist.shape
        ws.set_row(0, None, fmt_header)
        ws.set_column(0, ncol-1, None, fmt_base)
        for c in df_sizelist.columns:
            if np.issubdtype(df_sizelist[c].dtype, np.datetime64):
                ws.set_column(df_sizelist.columns.get_loc(c), df_sizelist.columns.get_loc(c), 12, fmt_date)
            elif c == "Order Quantity" or re.match(r'(?i)^UK_', c):
                ws.set_column(df_sizelist.columns.get_loc(c), df_sizelist.columns.get_loc(c), None, fmt_num)
        color_rows(ws, df_sizelist)
        autofit(ws, df_sizelist)
        ws.freeze_panes(1,0)

        # Subtotal sheets
        for name, df in [("Sizes", sizes_df), ("CRD_Mth_Sizes", crd_df), ("CRDPD_Mth_Sizes", crdpd_df)]:
            df.to_excel(writer, sheet_name=name, index=False)
            wsx = writer.sheets[name]
            wsx.set_row(0, None, fmt_header)
            color_rows(wsx, df)
            autofit(wsx, df)
            wsx.freeze_panes(1,0)

    return output.getvalue()

excel_bytes = build_excel_bytes(df_sizelist, sizes_df, crd_df, crdpd_df, cancel_sos)
st.download_button("⬇️ Download Excel (warna permanen)",
                   data=excel_bytes,
                   file_name="df_sizelist_ready.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("🔴 Merah = New Order  •  🟣 Ungu = Cancel SO (indikasi visual, remark tetap 'cfm')")
