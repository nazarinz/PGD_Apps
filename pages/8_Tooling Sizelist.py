# ==========================================
# 8_Tooling Sizelist.py — PGD Apps (v8 stable)
# ==========================================
import re
import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="PGD Apps — Tooling Sizelist", page_icon="📊", layout="wide")
st.title("📊 PGD Tooling Sizelist — Subtotal Generator (Final v8)")

# ================= Upload & Input =================
uploaded = st.file_uploader("📤 Upload file Excel (SAP/In-house Sizelist)", type=["xlsx", "xls"])
if uploaded is None:
    st.info("⬆️ Silakan upload file Excel terlebih dahulu sebelum melanjutkan.")
    st.stop()

# Panduan untuk user awam
st.markdown("""
### 📘 Format File Excel Wajib
Pastikan file memiliki kolom berikut:

| Kolom Wajib        | Keterangan                                                   |
|--------------------|--------------------------------------------------------------|
| **Sales Order**     | Nomor SO unik                                                |
| **Document Date**   | Tanggal dokumen order                                       |
| **Article**         | Kode artikel (gunakan prefix **FG** atau **HS** saja)       |
| **Order Quantity**  | Jumlah order                                                |
| **CRD**             | Customer Request Date                                       |
| **PD**              | Planned Date                                                |
| **LPD**             | Latest Planned Date                                         |
| **Working Status**  | Status pengerjaan (wajib untuk deteksi “New”)               |
| **UK_***            | Kolom ukuran (size breakdown)                               |
""")

# Baca Excel
df_sizelist = pd.read_excel(uploaded)

# Filter Article hanya FG / HS
if "Article" in df_sizelist.columns:
    df_sizelist = df_sizelist[df_sizelist["Article"].astype(str).str.startswith(("FG", "HS"))]

# Input tambahan dari user
st.subheader("⚙️ Pengaturan Eksekusi")
new_order_date = st.date_input("Tanggal New Order terakhir *wajib diisi*", value=None)
cancel_sos_input = st.text_area("Daftar Sales Order Cancel (pisahkan dengan koma):", placeholder="contoh: 10897552, 10896721")

if st.button("🚀 Execute Generate"):
    if not new_order_date:
        st.error("❌ Silakan isi tanggal New Order terlebih dahulu sebelum mengeksekusi.")
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
    ins_pos = df_sizelist.columns.get_loc("Document Date") + 1
    df_sizelist.insert(ins_pos, "Remark", remark)

    # ================= Isi LPD kosong =================
    if {"CRD", "PD", "LPD"}.issubset(df_sizelist.columns):
        row_min = pd.concat([df_sizelist["CRD"], df_sizelist["PD"]], axis=1).min(axis=1, skipna=True)
        df_sizelist.loc[df_sizelist["LPD"].isna(), "LPD"] = row_min[df_sizelist["LPD"].isna()]

    # ================= Helpers =================
    def insert_after(df, after_col, new_col, values):
        pos = df.columns.get_loc(after_col) + 1
        df.insert(pos, new_col, values)

    def bucket_day(day):
        return np.where(day >= 24, "30",
               np.where(day >= 16, "23",
               np.where(day >= 8, "15",
               np.where(day >= 1, "07", None))))

    def fmt_mdy(dt):
        dt = pd.to_datetime(dt)
        return f"{dt.month}/{dt.day}/{dt.year}"

    # ================= CRD_Mth & CRDPD_Mth =================
    if "CRD" in df_sizelist.columns:
        YM_CRD = df_sizelist["CRD"].dt.strftime("%Y%m")
        Day_CRD = df_sizelist["CRD"].dt.day
        base = pd.Series(bucket_day(Day_CRD), index=df_sizelist.index)
        Remark = df_sizelist["Remark"].str.lower().fillna("cfm")
        CRD_Mth = YM_CRD.fillna("") + base.fillna("") + "_" + Remark
        insert_after(df_sizelist, "CRD", "YM_CRD", YM_CRD)
        insert_after(df_sizelist, "YM_CRD", "Day_CRD", Day_CRD)
        insert_after(df_sizelist, "Day_CRD", "Class_CRD", base)
        insert_after(df_sizelist, "Class_CRD", "CRD_Mth", CRD_Mth)

    if "LPD" in df_sizelist.columns:
        YM_LPD = df_sizelist["LPD"].dt.strftime("%Y%m")
        Day_LPD = df_sizelist["LPD"].dt.day
        base2 = pd.Series(bucket_day(Day_LPD), index=df_sizelist.index)
        Remark = df_sizelist["Remark"].str.lower().fillna("cfm")
        CRDPD_Mth = YM_LPD.fillna("") + base2.fillna("") + "_" + Remark
        insert_after(df_sizelist, "LPD", "YM_CRDPD", YM_LPD)
        insert_after(df_sizelist, "YM_CRDPD", "Day_CRDPD", Day_LPD)
        insert_after(df_sizelist, "Day_CRDPD", "Class_CRDPD", base2)
        insert_after(df_sizelist, "Class_CRDPD", "CRDPD_Mth", CRDPD_Mth)

    # ================= Subtotal Builder =================
    size_cols = [c for c in df_sizelist.columns if re.match(r'(?i)^UK_', str(c))]
    order_cols = ["Order Quantity"] + size_cols

    def make_subtotal(df, group_col):
        pieces = []
        for key, grp in df.groupby(group_col, dropna=False):
            subtotal = {col: "" for col in ["Remark", group_col] + order_cols}
            subtotal["Remark"] = (
                "New" if (grp["Remark"].str.lower() == "new").any()
                else "cfm"
            )
            subtotal[group_col] = key
            for col in order_cols:
                subtotal[col] = grp[col].sum(skipna=True)
            pieces.append(subtotal)
        return pd.DataFrame(pieces)

    sizes_df = make_subtotal(df_sizelist, "Document Date")
    crd_df = make_subtotal(df_sizelist, "CRD_Mth")
    crdpd_df = make_subtotal(df_sizelist, "CRDPD_Mth")

    # ================= Pewarnaan =================
    def colorize(df):
        colors = []
        for _, row in df.iterrows():
            if str(row.get("Remark", "")).lower() == "new":
                colors.append("red")
            elif str(row.get("Sales Order", "")) in cancel_sos:
                colors.append("purple")
            else:
                colors.append("black")
        return colors

    color_main = colorize(df_sizelist)
    color_sizes = colorize(sizes_df)
    color_crd = colorize(crd_df)
    color_crdpd = colorize(crdpd_df)

    # ================= Preview =================
    st.subheader("📑 Data (preview)")
    st.dataframe(df_sizelist.head(20), use_container_width=True)

    # ================= Excel Export =================
    def build_excel():
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="m/d/yyyy") as writer:
            wb = writer.book
            fmt = lambda color=None: wb.add_format({
                "font_name": "Calibri", "font_size": 9,
                "align": "center", "valign": "vcenter",
                "font_color": color or "black", "num_format": "m/d/yyyy"
            })

            def write(ws, df, colors):
                for j, col in enumerate(df.columns):
                    ws.write(0, j, col, fmt("black").set_bold())
                for i, (_, row) in enumerate(df.iterrows()):
                    for j, val in enumerate(row):
                        ws.write(i + 1, j, val, fmt(colors[i]))
                ws.freeze_panes(1, 0)

            df_sizelist.to_excel(writer, sheet_name="Data", index=False)
            ws1 = writer.sheets["Data"]
            write(ws1, df_sizelist, color_main)

            sizes_df.to_excel(writer, sheet_name="Sizes", index=False)
            ws2 = writer.sheets["Sizes"]
            write(ws2, sizes_df, color_sizes)

            crd_df.to_excel(writer, sheet_name="CRD_Mth_Sizes", index=False)
            ws3 = writer.sheets["CRD_Mth_Sizes"]
            write(ws3, crd_df, color_crd)

            crdpd_df.to_excel(writer, sheet_name="CRDPD_Mth_Sizes", index=False)
            ws4 = writer.sheets["CRDPD_Mth_Sizes"]
            write(ws4, crdpd_df, color_crdpd)

            # Summary sheet
            total_so = df_sizelist["Sales Order"].nunique() if "Sales Order" in df_sizelist else 0
            total_new = (df_sizelist["Remark"].str.lower() == "new").sum()
            total_cancel = sum(df_sizelist["Sales Order"].astype(str).isin(cancel_sos))
            summary = pd.DataFrame({
                "Metric": ["Total SO", "Total New", "Total Cancel"],
                "Value": [total_so, total_new, total_cancel]
            })
            summary.to_excel(writer, sheet_name="Summary", index=False)

        return output.getvalue()

    excel_bytes = build_excel()
    st.download_button("⬇️ Download Excel", data=excel_bytes,
                       file_name="Tooling_Sizelist_v8.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
