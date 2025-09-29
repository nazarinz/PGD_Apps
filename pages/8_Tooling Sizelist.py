import re
import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="PGD Apps — Subtotal Generator", page_icon="📊", layout="wide")
st.title("📊 Subtotal Generator — Sizes, CRD_Mth, CRDPD_Mth")

# ================= Panduan Awal =================
st.subheader("📖 Panduan Penggunaan")
st.markdown("Sebelum mengunggah file Excel, pastikan file Anda memiliki kolom-kolom berikut:")

cols_info = pd.DataFrame({
    "Nama Kolom": [
        "Document Date", "LPD", "CRD", "PD", "Order Quantity", "Article",
        "UK_* (contoh: UK_3K, UK_4K, UK_5, dst.)"
    ],
    "Keterangan": [
        "Tanggal dokumen pesanan (format date)",
        "Latest Planned Delivery (boleh kosong, akan diisi otomatis)",
        "Customer Request Date (format date)",
        "Planned Date (format date)",
        "Jumlah order (angka)",
        "Kode artikel (hanya FG / HS yang dipakai, HL & HU diabaikan)",
        "Kolom ukuran, wajib prefix 'UK_' untuk setiap size"
    ]
})
st.table(cols_info)

# ================= Upload & Input =================
uploaded = st.file_uploader("⬆️ Upload file Excel", type=["xlsx", "xls"])
if uploaded is None:
    st.info("Silakan upload file Excel terlebih dahulu untuk melanjutkan.")
    st.stop()

new_order_date = st.date_input("📅 Tanggal New Order terakhir", pd.to_datetime("today"))
NEW_ORDER_DATE = pd.to_datetime(new_order_date)

# ================= Load data =================
df_sizelist = pd.read_excel(uploaded)

# Filter hanya Article yang depannya FG atau HS
if "Article" in df_sizelist.columns:
    df_sizelist = df_sizelist[df_sizelist["Article"].astype(str).str.startswith(("FG", "HS"))].copy()
    df_sizelist.reset_index(drop=True, inplace=True)

# ================= Normalisasi tanggal penting =================
for col in ["Document Date", "LPD", "CRD", "PD"]:
    if col in df_sizelist.columns:
        df_sizelist[col] = pd.to_datetime(df_sizelist[col], errors="coerce")

# ================= Remark =================
if "Remark" in df_sizelist.columns:
    df_sizelist.drop(columns=["Remark"], inplace=True)

remark = np.where(
    (df_sizelist["Document Date"] >= NEW_ORDER_DATE) & (df_sizelist["LPD"].isna()),
    "New", "cfm"
)
ins_pos = df_sizelist.columns.get_loc("Document Date") + 1
df_sizelist.insert(ins_pos, "Remark", remark)

# ================= LPD kosong -> min(CRD, PD) =================
if {"CRD","PD","LPD"}.issubset(df_sizelist.columns):
    row_min = pd.concat([df_sizelist["CRD"], df_sizelist["PD"]], axis=1).min(axis=1, skipna=True)
    df_sizelist.loc[df_sizelist["LPD"].isna(), "LPD"] = row_min[df_sizelist["LPD"].isna()]

# ================= Helpers =================
def insert_after(df, after_col, new_col, values):
    pos = df.columns.get_loc(after_col) + 1
    df.insert(pos, new_col, values)

def bucket_base_from_day(day: pd.Series) -> pd.Series:
    """Return bucket base: '30','23','15','07' or None per day."""
    return np.where(day >= 24, "30",
           np.where(day >= 16, "23",
           np.where(day >= 8,  "15",
           np.where(day >= 1,  "07", None))))

def fmt_mdy(dt):
    dt = pd.to_datetime(dt)
    return f"{dt.month}/{dt.day}/{dt.year}"

def make_subtotal_only(df_source, group_col, order_cols, label_fmt, remark_policy="any-new"):
    """
    Build dataframe subtotal-only per group, tanpa detail, + Grand Total.
    remark_policy:
      - 'any-new'      : subtotal Remark=New jika ada baris New dalam grup (dipakai untuk Document Date).
      - 'suffix-based' : subtotal Remark mengikuti akhiran key (_new=>New; selainnya=>cfm) (dipakai CRD/CRDPD).
    """
    if group_col not in df_source.columns:
        raise ValueError(f"Kolom grup '{group_col}' tidak ada di data sumber.")

    df_vis = df_source[[group_col] + order_cols].copy().sort_values(group_col, ascending=True, na_position="last")
    df_work = pd.concat(
        [df_source.loc[df_vis.index, ["Remark"]].reset_index(drop=True),
         df_vis.reset_index(drop=True)],
        axis=1
    )

    pieces = []
    for key, grp in df_work.groupby(df_work[group_col], dropna=False, sort=True):
        subtotal = {col: "" for col in df_work.columns}
        label = "(blank)" if pd.isna(key) else label_fmt(key)

        if remark_policy == "suffix-based":
            key_str = "" if pd.isna(key) else str(key)
            subtotal["Remark"] = "New" if key_str.endswith("_new") else "cfm"
        else:  # any-new
            has_new = (grp["Remark"].astype(str).str.lower() == "new").any()
            subtotal["Remark"] = "New" if has_new else "cfm"

        subtotal[group_col] = label
        for col in order_cols:
            subtotal[col] = grp[col].sum(skipna=True)
        pieces.append(pd.DataFrame([subtotal], columns=df_work.columns))

    out = pd.concat(pieces, ignore_index=True)

    grand_vals = df_vis.drop(columns=[group_col]).sum(numeric_only=True)
    grand = {col:"" for col in out.columns}
    grand[group_col] = "Grand Total"
    grand["Remark"]  = "cfm"
    for col in order_cols:
        grand[col] = grand_vals.get(col, "")
    out = pd.concat([out, pd.DataFrame([grand])], ignore_index=True)
    return out

# ================= Turunan CRD (grup-based, anti duplikat) =================
if "CRD" in df_sizelist.columns:
    YM_CRD  = df_sizelist["CRD"].dt.strftime("%Y%m")
    Day_CRD = df_sizelist["CRD"].dt.day
    BucketB = pd.Series(bucket_base_from_day(Day_CRD), index=df_sizelist.index, dtype="object")

    tmp = pd.DataFrame({"YM": YM_CRD, "BucketBase": BucketB, "Remark": df_sizelist["Remark"]})
    suffix_map = {}
    for (ym, bb), grp in tmp.groupby(["YM", "BucketBase"], dropna=False):
        if pd.isna(ym) or pd.isna(bb):
            continue
        has_new = grp["Remark"].astype(str).str.lower().eq("new").any()
        suffix_map[(ym, bb)] = "new" if has_new else "cfm"

    final_class_crd = []
    for ym, bb in zip(YM_CRD, BucketB):
        if pd.isna(ym) or pd.isna(bb):
            final_class_crd.append(None)
        else:
            suff = suffix_map.get((ym, bb), "cfm")
            final_class_crd.append(f"{bb}_{suff}")

    CRD_Mth = YM_CRD.fillna("").astype(str) + pd.Series(final_class_crd).fillna("")

    insert_after(df_sizelist, "CRD", "YM_CRD", YM_CRD)
    insert_after(df_sizelist, "YM_CRD", "Day_CRD", Day_CRD)
    insert_after(df_sizelist, "Day_CRD", "Class_CRD", pd.Series(final_class_crd, dtype="object"))
    insert_after(df_sizelist, "Class_CRD", "CRD_Mth", CRD_Mth)

# ================= Turunan LPD (grup-based, anti duplikat) =================
if "LPD" in df_sizelist.columns:
    YM_CRDPD  = df_sizelist["LPD"].dt.strftime("%Y%m")
    Day_CRDPD = df_sizelist["LPD"].dt.day
    BucketB2  = pd.Series(bucket_base_from_day(Day_CRDPD), index=df_sizelist.index, dtype="object")

    tmp2 = pd.DataFrame({"YM": YM_CRDPD, "BucketBase": BucketB2, "Remark": df_sizelist["Remark"]})
    suffix_map2 = {}
    for (ym, bb), grp in tmp2.groupby(["YM", "BucketBase"], dropna=False):
        if pd.isna(ym) or pd.isna(bb):
            continue
        has_new = grp["Remark"].astype(str).str.lower().eq("new").any()
        suffix_map2[(ym, bb)] = "new" if has_new else "cfm"

    final_class_crdpd = []
    for ym, bb in zip(YM_CRDPD, BucketB2):
        if pd.isna(ym) or pd.isna(bb):
            final_class_crdpd.append(None)
        else:
            suff = suffix_map2.get((ym, bb), "cfm")
            final_class_crdpd.append(f"{bb}_{suff}")

    CRDPD_Mth = YM_CRDPD.fillna("").astype(str) + pd.Series(final_class_crdpd).fillna("")

    insert_after(df_sizelist, "LPD", "YM_CRDPD", YM_CRDPD)
    insert_after(df_sizelist, "YM_CRDPD", "Day_CRDPD", Day_CRDPD)
    insert_after(df_sizelist, "Day_CRDPD", "Class_CRDPD", pd.Series(final_class_crdpd, dtype="object"))
    insert_after(df_sizelist, "Class_CRDPD", "CRDPD_Mth", CRDPD_Mth)

# ================= Subtotal Builder =================
size_cols = [c for c in df_sizelist.columns if re.match(r'(?i)^UK_', str(c))]
order_cols = ["Order Quantity"] + size_cols

# Document Date → subtotal Remark: any-new (sesuai sebelumnya)
sizes_df  = make_subtotal_only(
    df_sizelist, "Document Date", order_cols,
    label_fmt=lambda k: f"{fmt_mdy(k)} Total",
    remark_policy="any-new"
)

# CRD/CRDPD → subtotal Remark mengikuti suffix (_new => New, lainnya => cfm)
crd_df    = make_subtotal_only(
    df_sizelist, "CRD_Mth", order_cols,
    label_fmt=lambda k: str(k),
    remark_policy="suffix-based"
)

crdpd_df  = make_subtotal_only(
    df_sizelist, "CRDPD_Mth", order_cols,
    label_fmt=lambda k: str(k),
    remark_policy="suffix-based"
).rename(columns={"CRDPD_Mth": "CRDPD_month"})

# ================= Preview =================
st.subheader("📑 Data (preview)")
st.dataframe(df_sizelist.head(20), use_container_width=True)

c1, c2 = st.columns(2)
with c1:
    st.subheader("📑 Sizes (Subtotal Only)")
    st.dataframe(sizes_df.head(20), use_container_width=True)
with c2:
    st.subheader("📑 CRD_Mth_Sizes (Subtotal Only)")
    st.dataframe(crd_df.head(20), use_container_width=True)

st.subheader("📑 CRDPD_Mth_Sizes (Subtotal Only)")
st.dataframe(crdpd_df.head(20), use_container_width=True)

# ================= Sanity Check Totals =================
st.subheader("🔍 Sanity Check Totals")
totals = {
    "Data (detail)": pd.to_numeric(df_sizelist["Order Quantity"], errors="coerce").sum(skipna=True),
    "Sizes (subtotal)": pd.to_numeric(
        sizes_df.loc[sizes_df["Document Date"].eq("Grand Total"), "Order Quantity"], errors="coerce"
    ).sum(skipna=True),
    "CRD_Mth_Sizes (subtotal)": pd.to_numeric(
        crd_df.loc[crd_df["CRD_Mth"].eq("Grand Total"), "Order Quantity"], errors="coerce"
    ).sum(skipna=True),
    "CRDPD_Mth_Sizes (subtotal)": pd.to_numeric(
        crdpd_df.loc[crdpd_df["CRDPD_month"].eq("Grand Total"), "Order Quantity"], errors="coerce"
    ).sum(skipna=True),
}
check_df = pd.DataFrame(list(totals.items()), columns=["Sheet", "Grand Total Order Qty"])
st.table(check_df)

if len({round(v or 0, 6) for v in totals.values()}) == 1:
    st.success("✅ Semua Grand Total konsisten di semua sheet.")
else:
    st.error("⚠️ Grand Total tidak konsisten! Cek kembali logika subtotal.")

# ================= Export Excel (HARD-CODE RED) =================
def build_excel_bytes(df_sizelist, sizes_df, crd_df, crdpd_df) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="m/d/yyyy") as writer:
        wb = writer.book

        # Styles hitam
        base   = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter"})
        datef  = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","num_format":"m/d/yyyy"})
        header = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","bold":True})
        num0   = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","num_format":"0"})
        # Styles merah hard-coded
        red_base = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","font_color":"#FF0000"})
        red_date = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","font_color":"#FF0000","num_format":"m/d/yyyy"})
        red_num  = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","font_color":"#FF0000","num_format":"0"})

        def autofit(ws, df, skip_cols=[]):
            for j, col in enumerate(df.columns):
                if j in skip_cols:
                    continue
                maxlen = max(len(str(col)), df[col].astype(str).map(len).max() if len(df)>0 else 0)
                ws.set_column(j, j, max(8, min(60, maxlen + 2)))

        def write_cell(ws, row, col, val, fmt):
            if pd.isna(val):
                ws.write_blank(row, col, None, fmt)
            elif isinstance(val, (pd.Timestamp, datetime)):
                ws.write_datetime(row, col, pd.to_datetime(val).to_pydatetime(), fmt)
            elif isinstance(val, (int, float, np.integer, np.floating)) and pd.api.types.is_number(val):
                ws.write_number(row, col, float(val), fmt)
            else:
                ws.write_string(row, col, str(val), fmt)

        def repaint_red_rows(ws, df, red_row_mask, date_cols_idx, num_cols_idx, hide_first_col=False):
            start_row = 1  # data start after header
            for i, is_red in enumerate(red_row_mask):
                if not is_red:
                    continue
                excel_row = start_row + i
                for j in range(len(df.columns)):
                    if hide_first_col and j == 0:
                        continue
                    v = df.iat[i, j]
                    if j in date_cols_idx:
                        fmt = red_date
                    elif j in num_cols_idx:
                        fmt = red_num
                    else:
                        fmt = red_base
                    write_cell(ws, excel_row, j, v, fmt)

        # ---- SHEET Data ----
        df_sizelist.to_excel(writer, "Data", index=False)
        ws1 = writer.sheets["Data"]
        nrow1, ncol1 = df_sizelist.shape
        ws1.set_column(0, ncol1-1, None, base)
        ws1.set_row(0, None, header)

        dt_cols_1 = [c for c in df_sizelist.columns if np.issubdtype(df_sizelist[c].dtype, np.datetime64)]
        dt_idx_1  = set(df_sizelist.columns.get_loc(c) for c in dt_cols_1)
        num_idx_1 = set()
        if "Order Quantity" in df_sizelist.columns:
            num_idx_1.add(df_sizelist.columns.get_loc("Order Quantity"))
        for j,c in enumerate(df_sizelist.columns):
            if re.match(r'(?i)^UK_', str(c)):
                num_idx_1.add(j)
        for c in dt_cols_1:
            idx = df_sizelist.columns.get_loc(c)
            ws1.set_column(idx, idx, 12, datef)
        for idx in sorted(num_idx_1):
            ws1.set_column(idx, idx, None, num0)

        if "Remark" in df_sizelist.columns:
            red_mask_1 = df_sizelist["Remark"].astype(str).str.lower().eq("new").to_numpy()
            repaint_red_rows(ws1, df_sizelist, red_mask_1, dt_idx_1, num_idx_1, hide_first_col=False)

        ws1.freeze_panes(1, 0)
        ws1.autofilter(0, 0, nrow1, ncol1-1)
        autofit(ws1, df_sizelist)

        # ---- SHEET Sizes ----
        sizes_df.to_excel(writer, "Sizes", index=False)
        ws2 = writer.sheets["Sizes"]
        nrow2, ncol2 = sizes_df.shape
        ws2.set_row(0, None, header)
        ws2.set_column(0, 0, 0)   # hide Remark

        dt_idx_2  = set()         # label subtotal string, bukan datetime
        num_idx_2 = set()
        if "Order Quantity" in sizes_df.columns:
            num_idx_2.add(sizes_df.columns.get_loc("Order Quantity"))
        for j,c in enumerate(sizes_df.columns):
            if re.match(r'(?i)^UK_', str(c)): num_idx_2.add(j)

        red_mask_2 = sizes_df["Remark"].astype(str).str.lower().eq("new").to_numpy()
        repaint_red_rows(ws2, sizes_df, red_mask_2, dt_idx_2, num_idx_2, hide_first_col=True)

        ws2.freeze_panes(1, 1)
        ws2.autofilter(0, 1, nrow2, ncol2-1)
        autofit(ws2, sizes_df, skip_cols=[0])

        # ---- SHEET CRD_Mth_Sizes ----
        crd_df.to_excel(writer, "CRD_Mth_Sizes", index=False)
        ws3 = writer.sheets["CRD_Mth_Sizes"]
        nrow3, ncol3 = crd_df.shape
        ws3.set_row(0, None, header)
        ws3.set_column(0, 0, 0)   # hide Remark

        dt_idx_3, num_idx_3 = set(), set()
        if "Order Quantity" in crd_df.columns:
            num_idx_3.add(crd_df.columns.get_loc("Order Quantity"))
        for j,c in enumerate(crd_df.columns):
            if re.match(r'(?i)^UK_', str(c)): num_idx_3.add(j)

        red_mask_3 = crd_df["Remark"].astype(str).str.lower().eq("new").to_numpy()
        repaint_red_rows(ws3, crd_df, red_mask_3, dt_idx_3, num_idx_3, hide_first_col=True)

        ws3.freeze_panes(1, 1)
        ws3.autofilter(0, 1, nrow3, ncol3-1)
        autofit(ws3, crd_df, skip_cols=[0])

        # ---- SHEET CRDPD_Mth_Sizes ----
        crdpd_df.to_excel(writer, "CRDPD_Mth_Sizes", index=False)
        ws4 = writer.sheets["CRDPD_Mth_Sizes"]
        nrow4, ncol4 = crdpd_df.shape
        ws4.set_row(0, None, header)
        ws4.set_column(0, 0, 0)   # hide Remark

        dt_idx_4, num_idx_4 = set(), set()
        if "Order Quantity" in crdpd_df.columns:
            num_idx_4.add(crdpd_df.columns.get_loc("Order Quantity"))
        for j,c in enumerate(crdpd_df.columns):
            if re.match(r'(?i)^UK_', str(c)): num_idx_4.add(j)

        red_mask_4 = crdpd_df["Remark"].astype(str).str.lower().eq("new").to_numpy()
        repaint_red_rows(ws4, crdpd_df, red_mask_4, dt_idx_4, num_idx_4, hide_first_col=True)

        ws4.freeze_panes(1, 1)
        ws4.autofilter(0, 1, nrow4, ncol4-1)
        autofit(ws4, crdpd_df, skip_cols=[0])

    return output.getvalue()

excel_bytes = build_excel_bytes(df_sizelist, sizes_df, crd_df, crdpd_df)
st.download_button(
    "⬇️ Download Excel (match Colab, hard-coded red)",
    data=excel_bytes,
    file_name="df_sizelist_ready.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
