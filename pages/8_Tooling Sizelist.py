# ==========================================
# 8_Tooling Sizelist.py — PGD Apps (v10.2 hardcolor full-fix)
# ==========================================
import re
import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="PGD Apps — Tooling Sizelist", page_icon="📊", layout="wide")
st.title("📊 PGD Tooling Sizelist — Subtotal Generator (v10.2 Final)")

# ================= Upload =================
uploaded = st.file_uploader("📤 Upload file Excel (SAP/In-house Sizelist)", type=["xlsx", "xls"])
if uploaded is None:
    st.info("⬆️ Silakan upload file Excel terlebih dahulu sebelum melanjutkan.")
    st.stop()

# Panduan
st.markdown("""
### 📘 Format File Excel Wajib
| Kolom Wajib | Keterangan |
|--------------|------------|
| **Sales Order** | Nomor SO unik |
| **Document Date** | Tanggal dokumen order |
| **Article** | Gunakan prefix FG / HS saja |
| **Order Quantity** | Jumlah order |
| **CRD** | Customer Request Date |
| **PD** | Planned Date |
| **LPD** | Latest Planned Date |
| **Working Status** | Status pengerjaan (untuk deteksi “New”) |
| **UK_*** | Kolom ukuran (size breakdown) |
""")

df_sizelist = pd.read_excel(uploaded)

# Filter Article hanya FG / HS
if "Article" in df_sizelist.columns:
    df_sizelist = df_sizelist[df_sizelist["Article"].astype(str).str.startswith(("FG", "HS"))]

# Input user
st.subheader("⚙️ Pengaturan Eksekusi")
new_order_date = st.date_input("Tanggal New Order terakhir *wajib diisi*", value=None)
cancel_sos_input = st.text_area("Daftar Sales Order Cancel (pisahkan dengan koma):", placeholder="contoh: 10897552, 10896721")

if st.button("🚀 Execute Generate"):
    if not new_order_date:
        st.error("❌ Silakan isi tanggal New Order terlebih dahulu.")
        st.stop()

    NEW_ORDER_DATE = pd.to_datetime(new_order_date)
    cancel_sos = [s.strip() for s in cancel_sos_input.split(",") if s.strip()]

    # ================= Normalisasi tanggal =================
    for col in ["Document Date", "LPD", "CRD", "PD"]:
        if col in df_sizelist.columns:
            df_sizelist[col] = pd.to_datetime(df_sizelist[col], errors="coerce")

    # ================= Remark =================
    if "Remark" in df_sizelist.columns:
        df_sizelist.drop(columns=["Remark"], inplace=True)

    remark = np.where(
        (df_sizelist["Document Date"] >= NEW_ORDER_DATE)
        & (df_sizelist["LPD"].isna())
        & (df_sizelist.get("Working Status", "").astype(str).str.strip() == "10"),
        "New", "cfm"
    )
    df_sizelist.insert(df_sizelist.columns.get_loc("Document Date") + 1, "Remark", remark)

    # ================= Isi LPD kosong =================
    if {"CRD", "PD", "LPD"}.issubset(df_sizelist.columns):
        row_min = pd.concat([df_sizelist["CRD"], df_sizelist["PD"]], axis=1).min(axis=1, skipna=True)
        df_sizelist.loc[df_sizelist["LPD"].isna(), "LPD"] = row_min[df_sizelist["LPD"].isna()]

    # ================= CRD & CRDPD Mth =================
    def insert_after(df, after_col, new_col, values):
        pos = df.columns.get_loc(after_col) + 1
        df.insert(pos, new_col, values)

    def bucket_day(day):
        return np.where(day >= 24, "30",
               np.where(day >= 16, "23",
               np.where(day >= 8, "15",
               np.where(day >= 1, "07", None))))

    if "CRD" in df_sizelist.columns:
        YM = df_sizelist["CRD"].dt.strftime("%Y%m")
        Day = df_sizelist["CRD"].dt.day
        base = pd.Series(bucket_day(Day))
        CRD_Mth = YM.fillna("") + base.fillna("") + "_" + df_sizelist["Remark"].str.lower()
        insert_after(df_sizelist, "CRD", "CRD_Mth", CRD_Mth)

    if "LPD" in df_sizelist.columns:
        YM = df_sizelist["LPD"].dt.strftime("%Y%m")
        Day = df_sizelist["LPD"].dt.day
        base2 = pd.Series(bucket_day(Day))
        CRDPD_Mth = YM.fillna("") + base2.fillna("") + "_" + df_sizelist["Remark"].str.lower()
        insert_after(df_sizelist, "LPD", "CRDPD_Mth", CRDPD_Mth)

    # ================= Subtotal =================
    size_cols = [c for c in df_sizelist.columns if re.match(r'(?i)^UK_', str(c))]
    order_cols = ["Order Quantity"] + size_cols

    def make_subtotal(df, group_col):
        data, colors = [], []
        for key, grp in df.groupby(group_col, dropna=False):
            subtotal = {group_col: key}
            for c in order_cols:
                subtotal[c] = grp[c].sum(skipna=True)
            has_new = (grp["Remark"].str.lower() == "new").any()
            has_cancel = any(grp["Sales Order"].astype(str).isin(cancel_sos))
            color = "red" if has_new else ("purple" if has_cancel else "black")
            data.append(subtotal)
            colors.append(color)
        out = pd.DataFrame(data).reset_index(drop=True)
        return out, colors

    sizes_df, color_sizes = make_subtotal(df_sizelist, "Document Date")
    crd_df, color_crd = make_subtotal(df_sizelist, "CRD_Mth")
    crdpd_df, color_crdpd = make_subtotal(df_sizelist, "CRDPD_Mth")

    # Pewarnaan Data utama
    def colorize(df):
        colors = []
        for _, row in df.iterrows():
            so = str(row.get("Sales Order", ""))
            if so in cancel_sos:
                colors.append("purple")
            elif str(row.get("Remark", "")).lower() == "new":
                colors.append("red")
            else:
                colors.append("black")
        return colors

    color_main = colorize(df_sizelist)

    # ================= Excel Export =================
    def build_excel():
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="m/d/yyyy") as writer:
            wb = writer.book

            # Define solid formats (hard color)
            fmt_red = wb.add_format({"font_name": "Calibri", "font_size": 9,
                                     "align": "center", "valign": "vcenter",
                                     "font_color": "#FF0000"})
            fmt_purple = wb.add_format({"font_name": "Calibri", "font_size": 9,
                                        "align": "center", "valign": "vcenter",
                                        "font_color": "#7030A0"})
            fmt_black = wb.add_format({"font_name": "Calibri", "font_size": 9,
                                       "align": "center", "valign": "vcenter",
                                       "font_color": "#000000"})
            fmts = {"red": fmt_red, "purple": fmt_purple, "black": fmt_black}

            def write(ws, df, colors):
                for j, col in enumerate(df.columns):
                    ws.write(0, j, col, fmt_black)
                for i, (_, row) in enumerate(df.iterrows()):
                    fmt = fmts.get(colors[i] if i < len(colors) else "black", fmt_black)
                    for j, val in enumerate(row):
                        ws.write(i + 1, j, val if not pd.isna(val) else "", fmt)
                ws.freeze_panes(1, 0)
                ws.autofilter(0, 0, len(df), len(df.columns) - 1)

            # Data
            df_sizelist.to_excel(writer, sheet_name="Data", index=False)
            write(writer.sheets["Data"], df_sizelist, color_main)

            # Subtotals
            sizes_df.to_excel(writer, sheet_name="Sizes", index=False)
            write(writer.sheets["Sizes"], sizes_df, color_sizes)
            crd_df.to_excel(writer, sheet_name="CRD_Mth_Sizes", index=False)
            write(writer.sheets["CRD_Mth_Sizes"], crd_df, color_crd)
            crdpd_df.to_excel(writer, sheet_name="CRDPD_Mth_Sizes", index=False)
            write(writer.sheets["CRDPD_Mth_Sizes"], crdpd_df, color_crdpd)

        return output.getvalue()

    excel_bytes = build_excel()
    st.success("✅ Proses selesai dan warna aktif di semua sheet!")
    st.download_button("⬇️ Download Excel", data=excel_bytes,
                       file_name="Tooling_Sizelist_v10.2.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
