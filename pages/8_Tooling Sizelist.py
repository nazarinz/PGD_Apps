# app.py — PGD Subtotal Generator (FINAL v7 — fix warna hilang setelah drop kolom)
# Author: Nazarudin Zaini

import re
import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO

# ====================== Streamlit Config ======================
st.set_page_config(page_title="PGD Apps — Sizelist untuk Tooling", page_icon="📊", layout="wide")
st.title("📊 Subtotal Generator — Sizes, CRD_Mth, CRDPD_Mth")

# ====================== Upload Excel ======================
uploaded = st.file_uploader("📂 Upload file Excel", type=["xlsx", "xls"])
if uploaded is None:
    st.info("⬆️ Silakan upload file Excel terlebih dahulu")
    st.stop()

# Panduan
st.markdown("""
### 📑 Format File Excel (Minimal)
| Kolom | Keterangan |
|---|---|
| **Sales Order** | Nomor SO |
| **Document Date** | Tanggal dokumen order |
| **Article** | Hanya FG/HS (HL & HU dihapus) |
| **Order Quantity** | Jumlah pesanan |
| **Working Status** | Status (gunakan 10 untuk new order) |
| **CRD**, **PD**, **LPD** | Tanggal penting |
| **UK_*** | Kolom ukuran |
""")

df_sizelist = pd.read_excel(uploaded)

# Hapus Article HL/HU
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

ws_num = pd.to_numeric(df_sizelist.get("Working Status", pd.Series(index=df_sizelist.index)), errors="coerce")
remark = np.where(
    (df_sizelist["Document Date"] >= NEW_ORDER_DATE)
    & (df_sizelist["LPD"].isna())
    & (ws_num == 10),
    "New",
    "cfm"
)
insert_pos = df_sizelist.columns.get_loc("Document Date") + 1
df_sizelist.insert(insert_pos, "Remark", remark)

# ====================== Isi LPD kosong ======================
if {"CRD", "PD", "LPD"}.issubset(df_sizelist.columns):
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

# ====================== Subtotal builder ======================
size_cols = [c for c in df_sizelist.columns if re.match(r'(?i)^UK_', str(c))]
order_cols = ["Order Quantity"] + size_cols

def make_subtotal_only(df_source, group_col, order_cols, label_fmt, cancel_sos):
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
        has_new = (grp["Remark"].astype(str).str.lower() == "new").any()
        has_cancel = grp["Sales Order"].astype(str).isin(cancel_sos).any()
        if has_new:
            subtotal["Remark"] = "New"
        elif has_cancel:
            subtotal["Remark"] = "Cancel"
        else:
            subtotal["Remark"] = "cfm"
        subtotal["Sales Order"] = ""
        subtotal[group_col] = label
        for col in order_cols:
            subtotal[col] = grp[col].sum(skipna=True)
        pieces.append(pd.DataFrame([subtotal], columns=df_work.columns))

    out = pd.concat(pieces, ignore_index=True)
    grand_vals = df_vis.drop(columns=[group_col]).sum(numeric_only=True)
    grand = {col: "" for col in out.columns}
    grand[group_col] = "Grand Total"
    grand["Remark"] = "cfm"
    for col in order_cols:
        grand[col] = grand_vals.get(col, "")
    out = pd.concat([out, pd.DataFrame([grand])], ignore_index=True)
    return out

sizes_df  = make_subtotal_only(df_sizelist, "Document Date", order_cols, lambda k: f"{fmt_mdy(k)} Total", cancel_sos)
crd_df    = make_subtotal_only(df_sizelist, "CRD_Mth", order_cols, str, cancel_sos)
crdpd_df  = make_subtotal_only(df_sizelist, "CRDPD_Mth", order_cols, str, cancel_sos)
crdpd_df  = crdpd_df.rename(columns={"CRDPD_Mth": "CRDPD_month"})

# Simpan versi untuk warna (dengan Remark)
sizes_df_color = sizes_df.copy()
crd_df_color = crd_df.copy()
crdpd_df_color = crdpd_df.copy()

# Drop kolom remark & sales order dari versi final (untuk Excel)
for df in [sizes_df, crd_df, crdpd_df]:
    for col in ["Remark", "Sales Order"]:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)

# ====================== Preview ======================
st.success("✅ Data berhasil diproses!")
st.dataframe(df_sizelist.head(20), use_container_width=True)

# ====================== Excel Export ======================
def build_excel_bytes(df_sizelist, sizes_df, crd_df, crdpd_df, sizes_df_color, crd_df_color, crdpd_df_color, cancel_sos):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="m/d/yyyy") as writer:
        wb = writer.book
        fmt_header = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "font_name": "Calibri", "font_size": 9})

        def get_fmt(color=None):
            base = {"align": "center", "valign": "vcenter", "font_name": "Calibri", "font_size": 9}
            if color:
                base["font_color"] = color
            return wb.add_format(base)

        fmt_black = get_fmt()
        fmt_red = get_fmt("#FF0000")
        fmt_purple = get_fmt("#800080")

        def autofit(ws, df):
            for j, c in enumerate(df.columns):
                maxlen = max(len(str(c)), df[c].astype(str).map(len).max() if len(df) > 0 else 0)
                ws.set_column(j, j, max(8, min(60, maxlen + 2)))

        def write_colored(ws, df, color_ref, is_data=False):
            for i in range(len(df)):
                remark = str(color_ref.iloc[i].get("Remark", "")).lower() if "Remark" in color_ref.columns else ""
                so = str(color_ref.iloc[i].get("Sales Order", "")) if "Sales Order" in color_ref.columns else ""
                if remark == "new":
                    fmt = fmt_red
                elif is_data and so in cancel_sos:
                    fmt = fmt_purple
                elif remark == "cancel":
                    fmt = fmt_purple
                else:
                    fmt = fmt_black
                for j, val in enumerate(df.iloc[i]):
                    if pd.isna(val):
                        ws.write_blank(i + 1, j, None, fmt)
                    else:
                        ws.write(i + 1, j, val, fmt)

        # tulis semua sheet
        sheets = [
            ("Data", df_sizelist, df_sizelist, True),
            ("Sizes", sizes_df, sizes_df_color, False),
            ("CRD_Mth_Sizes", crd_df, crd_df_color, False),
            ("CRDPD_Mth_Sizes", crdpd_df, crdpd_df_color, False)
        ]
        for name, df, df_color, is_data in sheets:
            df.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            ws.set_row(0, None, fmt_header)
            write_colored(ws, df, df_color, is_data=is_data)
            autofit(ws, df)
            ws.freeze_panes(1, 0)

    return output.getvalue()

excel_bytes = build_excel_bytes(
    df_sizelist, sizes_df, crd_df, crdpd_df,
    sizes_df_color, crd_df_color, crdpd_df_color, cancel_sos
)

st.download_button("⬇️ Download Excel (warna aktif & bersih)",
                   data=excel_bytes,
                   file_name="df_sizelist_ready.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("🔴 Merah = New • 🟣 Ungu = Cancel")
