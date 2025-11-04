# ==========================================
# 8_Tooling Sizelist.py — PGD Apps (v12.0 Production)
# ==========================================
import re
import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="PGD Apps — Tooling Sizelist", page_icon="📊", layout="wide")
st.title("📊 PGD Tooling Sizelist — Subtotal Generator (v12.0 Production Ready)")

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

# Baca Excel
df_sizelist = pd.read_excel(uploaded)

# 🧹 Normalisasi nama kolom
df_sizelist.columns = df_sizelist.columns.str.strip().str.lower()

# Filter Article hanya FG / HS
if "article" in df_sizelist.columns:
    df_sizelist = df_sizelist[df_sizelist["article"].astype(str).str.startswith(("fg", "hs"), na=False)]

# ================= Input user =================
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
    for col in ["document date", "lpd", "crd", "pd"]:
        if col in df_sizelist.columns:
            df_sizelist[col] = pd.to_datetime(df_sizelist[col], errors="coerce")

    # ================= Tambah Remark =================
    if "remark" in df_sizelist.columns:
        df_sizelist.drop(columns=["remark"], inplace=True)

    # Kolom wajib aman
    for c in ["working status", "document date", "lpd"]:
        if c not in df_sizelist.columns:
            df_sizelist[c] = np.nan

    # Normalisasi Working Status agar bisa deteksi “10” dalam bentuk apapun
    work_status = df_sizelist["working status"].astype(str).str.strip().str.replace(".0", "", regex=False)

    remark = np.where(
        (pd.to_datetime(df_sizelist["document date"], errors="coerce") >= NEW_ORDER_DATE)
        & (df_sizelist["lpd"].isna() | (df_sizelist["lpd"].astype(str).str.strip() == ""))
        & (work_status == "10"),
        "New", "cfm"
    )
    insert_pos = df_sizelist.columns.get_loc("document date") + 1
    df_sizelist.insert(insert_pos, "remark", remark)

    # ================= Isi LPD kosong =================
    if {"crd", "pd", "lpd"}.issubset(df_sizelist.columns):
        row_min = pd.concat([df_sizelist["crd"], df_sizelist["pd"]], axis=1).min(axis=1, skipna=True)
        df_sizelist.loc[df_sizelist["lpd"].isna(), "lpd"] = row_min[df_sizelist["lpd"].isna()]

    # ================= CRD & CRDPD Mth =================
    def bucket_day(day):
        return np.where(day >= 24, "30",
               np.where(day >= 16, "23",
               np.where(day >= 8, "15",
               np.where(day >= 1, "07", None))))

    if "crd" in df_sizelist.columns:
        YM = df_sizelist["crd"].dt.strftime("%Y%m")
        Day = df_sizelist["crd"].dt.day
        base = pd.Series(bucket_day(Day))
        CRD_Mth = YM.fillna("") + base.fillna("") + "_" + df_sizelist["remark"].str.lower()
        df_sizelist["crd_mth"] = CRD_Mth

    if "lpd" in df_sizelist.columns:
        YM = df_sizelist["lpd"].dt.strftime("%Y%m")
        Day = df_sizelist["lpd"].dt.day
        base2 = pd.Series(bucket_day(Day))
        CRDPD_Mth = YM.fillna("") + base2.fillna("") + "_" + df_sizelist["remark"].str.lower()
        df_sizelist["crdpd_mth"] = CRDPD_Mth

    # ================= Subtotal =================
    size_cols = [c for c in df_sizelist.columns if re.match(r'(?i)^uk_', str(c))]
    order_cols = ["order quantity"] + size_cols

    def make_subtotal(df, group_col):
        data, colors = [], []
        for key, grp in df.groupby(group_col, dropna=False):
            subtotal = {group_col: key}
            for c in order_cols:
                subtotal[c] = grp[c].sum(skipna=True)
            has_new = (grp["remark"].str.lower() == "new").any()
            has_cancel = any(grp["sales order"].astype(str).isin(cancel_sos)) if "sales order" in grp.columns else False
            color = "red" if has_new else ("purple" if has_cancel else "black")
            data.append(subtotal)
            colors.append(color)
        out = pd.DataFrame(data).reset_index(drop=True)
        return out, colors

    sizes_df, color_sizes = make_subtotal(df_sizelist, "document date")
    crd_df, color_crd = make_subtotal(df_sizelist, "crd_mth")
    crdpd_df, color_crdpd = make_subtotal(df_sizelist, "crdpd_mth")

    # ================= Pewarnaan data utama =================
    def colorize(df):
        colors = []
        for _, row in df.iterrows():
            so = str(row.get("sales order", ""))
            if so in cancel_sos:
                colors.append("purple")
            elif str(row.get("remark", "")).lower() == "new":
                colors.append("red")
            else:
                colors.append("black")
        return colors

    color_main = colorize(df_sizelist)

    # ================= Excel Export =================
    def build_excel():
        output = BytesIO()
        import xlsxwriter
        wb = xlsxwriter.Workbook(output, {'in_memory': True})

        # Format warna
        fmt_red = wb.add_format({"font_name": "Calibri", "font_size": 9, "align": "center", "valign": "vcenter", "font_color": "#FF0000"})
        fmt_purple = wb.add_format({"font_name": "Calibri", "font_size": 9, "align": "center", "valign": "vcenter", "font_color": "#7030A0"})
        fmt_black = wb.add_format({"font_name": "Calibri", "font_size": 9, "align": "center", "valign": "vcenter", "font_color": "#000000"})
        fmts = {"red": fmt_red, "purple": fmt_purple, "black": fmt_black}

        def write_sheet(sheet_name, df, colors):
            ws = wb.add_worksheet(sheet_name)
            for j, col in enumerate(df.columns):
                ws.write(0, j, col, fmt_black)
            for i, (_, row) in enumerate(df.iterrows()):
                fmt = fmts.get(colors[i] if i < len(colors) else "black", fmt_black)
                for j, val in enumerate(row):
                    if pd.isna(val): val = ""
                    # Deteksi tanggal → tulis format shortdate
                    if isinstance(val, (pd.Timestamp, np.datetime64)):
                        ws.write_datetime(i + 1, j, pd.to_datetime(val).to_pydatetime(), fmt)
                    else:
                        ws.write(i + 1, j, str(val), fmt)
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, len(df), len(df.columns) - 1)
            ws.set_column(0, len(df.columns) - 1, 14)

        write_sheet("Data", df_sizelist, color_main)
        write_sheet("Sizes", sizes_df, color_sizes)
        write_sheet("CRD_Mth_Sizes", crd_df, color_crd)
        write_sheet("CRDPD_Mth_Sizes", crdpd_df, color_crdpd)

        wb.close()
        output.seek(0)
        return output.getvalue()

    excel_bytes = build_excel()
    st.success("✅ Semua sheet sudah berwarna dan deteksi 'New' robust!")
    st.download_button("⬇️ Download Excel", data=excel_bytes,
                       file_name="Tooling_Sizelist_v12.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
