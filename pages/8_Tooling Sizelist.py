# app.py — PGD Subtotal Generator (FINAL-STABLE)
# Author: Nazarudin Zaini

import re
import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO

# =============== Streamlit Config ===============
st.set_page_config(page_title="PGD Apps — Subtotal Generator", page_icon="📊", layout="wide")
st.title("📊 Subtotal Generator — Sizes, CRD_Mth, CRDPD_Mth")

# =============== Upload Excel ===============
uploaded = st.file_uploader("Upload file Excel", type=["xlsx", "xls"])
if uploaded is None:
    st.info("⬆️ Silakan upload file Excel terlebih dahulu")
    st.stop()

# =============== Panduan User Awam ===============
st.markdown("""
### 📑 Format File Excel yang Diharapkan
Pastikan file Excel memiliki kolom berikut (minimal):

| Kolom Wajib        | Keterangan |
|--------------------|------------|
| **Sales Order**      | Nomor Sales Order (unik per order) |
| **Document Date**    | Tanggal dokumen order |
| **Article**          | Kode artikel (hanya FG / HS digunakan, HL & HU dihapus) |
| **Order Quantity**   | Jumlah pesanan |
| **Working Status**   | Status produksi (gunakan untuk filter `10` pada order baru) |
| **CRD**              | Customer Request Date |
| **PD**               | Planned Date |
| **LPD**              | Latest Planned Date |
| **UK_***             | Kolom ukuran (size breakdown) |

Kolom lain boleh ada, tapi yang di atas wajib untuk kalkulasi subtotal.
""")

# =============== Baca Data ===============
df_sizelist = pd.read_excel(uploaded)

# Filter: hanya Article yang diawali FG/HS
if "Article" in df_sizelist.columns:
    df_sizelist = df_sizelist[~df_sizelist["Article"].astype(str).str.startswith(("HL", "HU"))]

# =============== Input Wajib ===============
new_order_date = st.date_input("📅 Pilih tanggal New Order terakhir (WAJIB)")
if not new_order_date:
    st.warning("❗ Harap pilih tanggal New Order terlebih dahulu.")
    st.stop()

cancel_so_text = st.text_area("Masukkan daftar Sales Order yang **Cancel** (pisahkan dengan koma atau enter):")
cancel_sos = re.split(r"[,\s]+", cancel_so_text.strip()) if cancel_so_text else []
cancel_sos = [s.strip() for s in cancel_sos if s.strip()]

NEW_ORDER_DATE = pd.to_datetime(new_order_date)

# =============== Normalisasi tanggal ===============
for col in ["Document Date", "LPD", "CRD", "PD"]:
    if col in df_sizelist.columns:
        df_sizelist[col] = pd.to_datetime(df_sizelist[col], errors="coerce")

# =============== Logic Remark (2 kondisi utama) ===============
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

# =============== Isi LPD kosong (min dari CRD & PD) ===============
if {"CRD","PD","LPD"}.issubset(df_sizelist.columns):
    row_min = pd.concat([df_sizelist["CRD"], df_sizelist["PD"]], axis=1).min(axis=1, skipna=True)
    df_sizelist.loc[df_sizelist["LPD"].isna(), "LPD"] = row_min[df_sizelist["LPD"].isna()]

# =============== Helpers ===============
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

# =============== CRD_Mth & CRDPD_Mth ===============
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

# =============== Subtotal Builder ===============
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

# =============== Preview ===============
st.subheader("📑 Data (Preview)")
st.dataframe(df_sizelist.head(20), use_container_width=True)

c1, c2 = st.columns(2)
with c1:
    st.subheader("📊 Sizes (Subtotal Only)")
    st.dataframe(sizes_df.head(20), use_container_width=True)
with c2:
    st.subheader("📊 CRD_Mth_Sizes (Subtotal Only)")
    st.dataframe(crd_df.head(20), use_container_width=True)
st.subheader("📊 CRDPD_Mth_Sizes (Subtotal Only)")
st.dataframe(crdpd_df.head(20), use_container_width=True)

# =============== Export to Excel ===============
def build_excel_bytes(df_sizelist, sizes_df, crd_df, crdpd_df, cancel_sos) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="m/d/yyyy") as writer:
        wb = writer.book
        base   = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter"})
        datef  = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","num_format":"m/d/yyyy"})
        header = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","bold":True})
        red    = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","font_color":"#FF0000"})
        purple = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","font_color":"#800080"})
        num0   = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","num_format":"0"})

        def excel_col(i):
            s=""; n=i+1
            while n: n, r = divmod(n-1, 26); s = chr(65+r)+s
            return s

        def autofit(ws, df, skip_cols=[]):
            for j, col in enumerate(df.columns):
                if j in skip_cols: continue
                maxlen = max(len(str(col)), df[col].astype(str).map(len).max() if len(df)>0 else 0)
                ws.set_column(j, j, max(8, min(60, maxlen + 2)))

        def apply_formats(ws, df, hide_first=False):
            nrow, ncol = df.shape
            ws.set_row(0, None, header)
            if hide_first: ws.set_column(0, 0, 0)
            last = excel_col(ncol-1)
            # warna merah utk New
            ws.conditional_format(f"A2:{last}{nrow+1}", {"type":"formula","criteria":'=INDIRECT("$A"&ROW())="New"',"format":red})
            ws.freeze_panes(1,1)
            ws.autofilter(0,1 if hide_first else 0,nrow,ncol-1)
            autofit(ws, df, skip_cols=[0] if hide_first else [])

        # === Sheet Data ===
        df_sizelist.to_excel(writer, sheet_name="Data", index=False)
        ws1 = writer.sheets["Data"]
        nrow1, ncol1 = df_sizelist.shape
        ws1.set_row(0, None, header)
        ws1.set_column(0, ncol1-1, None, base)
        dt_cols = [c for c in df_sizelist.columns if np.issubdtype(df_sizelist[c].dtype, np.datetime64)]
        for c in dt_cols:
            idx = df_sizelist.columns.get_loc(c)
            ws1.set_column(idx, idx, 12, datef)
        num_cols = [df_sizelist.columns.get_loc(c) for c in df_sizelist.columns if c == "Order Quantity" or re.match(r'(?i)^UK_', c)]
        for idx in sorted(set(num_cols)): ws1.set_column(idx, idx, None, num0)
        apply_formats(ws1, df_sizelist)
        # warna ungu utk SO Cancel
        if "Sales Order" in df_sizelist.columns and cancel_sos:
            so_idx = df_sizelist.columns.get_loc("Sales Order")
            so_col = excel_col(so_idx)
            last = excel_col(ncol1-1)
            for so in cancel_sos:
                ws1.conditional_format(
                    f"A2:{last}{nrow1+1}",
                    {"type":"formula","criteria":f'=INDIRECT("${so_col}"&ROW())="{so}"',"format":purple}
                )

        # === Sheet Sizes ===
        sizes_df.to_excel(writer, sheet_name="Sizes", index=False)
        apply_formats(writer.sheets["Sizes"], sizes_df, hide_first=True)

        # === Sheet CRD_Mth_Sizes ===
        crd_df.to_excel(writer, sheet_name="CRD_Mth_Sizes", index=False)
        apply_formats(writer.sheets["CRD_Mth_Sizes"], crd_df, hide_first=True)

        # === Sheet CRDPD_Mth_Sizes ===
        crdpd_df.to_excel(writer, sheet_name="CRDPD_Mth_Sizes", index=False)
        apply_formats(writer.sheets["CRDPD_Mth_Sizes"], crdpd_df, hide_first=True)

    return output.getvalue()

# =============== Download Button ===============
excel_bytes = build_excel_bytes(df_sizelist, sizes_df, crd_df, crdpd_df, cancel_sos)
st.download_button("⬇️ Download Excel (match manual Excel)", data=excel_bytes,
                   file_name="df_sizelist_ready.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
