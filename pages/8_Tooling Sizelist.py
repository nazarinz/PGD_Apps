import re
import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="PGD Apps — Subtotal Generator", page_icon="📊", layout="wide")
st.title("📊 Subtotal Generator — Sizes, CRD_Mth, CRDPD_Mth")

# ================= Panduan Awal =================
st.subheader("📖 Panduan Penggunaan")
st.markdown("""
Sebelum mengunggah file Excel, pastikan file Anda memiliki kolom-kolom berikut:
""")

cols_info = pd.DataFrame({
    "Nama Kolom": [
        "Document Date", 
        "LPD", 
        "CRD", 
        "PD", 
        "Order Quantity", 
        "Article", 
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

# Input tanggal New Order terakhir
new_order_date = st.date_input("📅 Tanggal New Order terakhir", pd.to_datetime("today"))
NEW_ORDER_DATE = pd.to_datetime(new_order_date)

# ================= Load data =================
df_sizelist = pd.read_excel(uploaded)

# Filter hanya Article yang depannya FG atau HS
if "Article" in df_sizelist.columns:
    df_sizelist = df_sizelist[df_sizelist["Article"].astype(str).str.startswith(("FG", "HS"))]

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

def day_to_class(day_series, suffix):
    return np.where(day_series >= 24, f"30_{suffix}",
           np.where(day_series >= 16, f"23_{suffix}",
           np.where(day_series >= 8,  f"15_{suffix}",
           np.where(day_series >= 1,  f"07_{suffix}", None))))

def fmt_mdy(dt):
    dt = pd.to_datetime(dt)
    return f"{dt.month}/{dt.day}/{dt.year}"

# ================= Turunan CRD =================
if "CRD" in df_sizelist.columns:
    YM_CRD  = df_sizelist["CRD"].dt.strftime("%Y%m")
    Day_CRD = df_sizelist["CRD"].dt.day
    is_new  = df_sizelist["Remark"].astype(str).str.lower().eq("new")
    Class_CRD = np.where(is_new, day_to_class(Day_CRD,"new"), day_to_class(Day_CRD,"cfm"))
    CRD_Mth = YM_CRD.fillna("").astype(str) + pd.Series(Class_CRD).fillna("")
    insert_after(df_sizelist, "CRD", "YM_CRD", YM_CRD)
    insert_after(df_sizelist, "YM_CRD", "Day_CRD", Day_CRD)
    insert_after(df_sizelist, "Day_CRD", "Class_CRD", Class_CRD)
    insert_after(df_sizelist, "Class_CRD", "CRD_Mth", CRD_Mth)

# ================= Turunan LPD =================
if "LPD" in df_sizelist.columns:
    YM_CRDPD  = df_sizelist["LPD"].dt.strftime("%Y%m")
    Day_CRDPD = df_sizelist["LPD"].dt.day
    suffix    = np.where(df_sizelist["Remark"].str.lower()=="new", "new", "cfm")
    Class_CRDPD = np.where(suffix=="new", day_to_class(Day_CRDPD,"new"), day_to_class(Day_CRDPD,"cfm"))
    CRDPD_Mth = YM_CRDPD.fillna("").astype(str) + pd.Series(Class_CRDPD).fillna("")
    insert_after(df_sizelist, "LPD", "YM_CRDPD", YM_CRDPD)
    insert_after(df_sizelist, "YM_CRDPD", "Day_CRDPD", Day_CRDPD)
    insert_after(df_sizelist, "Day_CRDPD", "Class_CRDPD", Class_CRDPD)
    insert_after(df_sizelist, "Class_CRDPD", "CRDPD_Mth", CRDPD_Mth)

# ================= Subtotal Builder =================
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
    # Grand Total
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

# ================= Excel Export =================
def build_excel_bytes(df_sizelist, sizes_df, crd_df, crdpd_df) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="m/d/yyyy") as writer:
        wb = writer.book
        base   = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter"})
        datef  = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","num_format":"m/d/yyyy"})
        header = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","bold":True})
        red    = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","font_color":"#FF0000"})
        num0   = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","num_format":"0"})

        def excel_col(i):
            s=""; n=i+1
            while n: n,r=divmod(n-1,26); s=chr(65+r)+s
            return s

        def autofit(ws, df, skip_cols=[]):
            for j, col in enumerate(df.columns):
                if j in skip_cols: continue
                maxlen = max(len(str(col)), df[col].astype(str).map(len).max() if len(df)>0 else 0)
                ws.set_column(j, j, max(8, min(60, maxlen + 2)))

        # Data sheet
        df_sizelist.to_excel(writer,"Data",index=False)
        ws1=writer.sheets["Data"]; nrow1,ncol1=df_sizelist.shape
        ws1.set_column(0,ncol1-1,None,base); ws1.set_row(0,None,header)
        for c in [c for c in df_sizelist.columns if np.issubdtype(df_sizelist[c].dtype,np.datetime64)]:
            idx=df_sizelist.columns.get_loc(c); ws1.set_column(idx,idx,12,datef)
        num_cols=[]; 
        if "Order Quantity" in df_sizelist.columns: num_cols.append(df_sizelist.columns.get_loc("Order Quantity"))
        for j,c in enumerate(df_sizelist.columns):
            if re.match(r'(?i)^UK_',str(c)): num_cols.append(j)
        for idx in sorted(set(num_cols)): ws1.set_column(idx,idx,None,num0)
        if "Remark" in df_sizelist.columns:
            last=excel_col(ncol1-1); rng=f"A2:{last}{nrow1+1}"
            ridx=df_sizelist.columns.get_loc("Remark"); rcol=excel_col(ridx)
            ws1.conditional_format(rng,{"type":"formula","criteria":f'=INDIRECT("${rcol}" & ROW())="New"',"format":red})
        ws1.freeze_panes(1,0); ws1.autofilter(0,0,nrow1,ncol1-1); autofit(ws1,df_sizelist)

        # Sizes
        sizes_df.to_excel(writer,"Sizes",index=False)
        ws2=writer.sheets["Sizes"]; nrow2,ncol2=sizes_df.shape
        ws2.set_row(0,None,header); ws2.set_column(0,0,0)
        last2=excel_col(ncol2-1)
        ws2.conditional_format(f"A2:{last2}{nrow2+1}",{"type":"formula","criteria":'=INDIRECT("$A"&ROW())="New"',"format":red})
        ws2.freeze_panes(1,1); ws2.autofilter(0,1,nrow2,ncol2-1); autofit(ws2,sizes_df,skip_cols=[0])

        # CRD_Mth_Sizes
        crd_df.to_excel(writer,"CRD_Mth_Sizes",index=False)
        ws3=writer.sheets["CRD_Mth_Sizes"]; nrow3,ncol3=crd_df.shape
        ws3.set_row(0,None,header); ws3.set_column(0,0,0)
        last3=excel_col(ncol3-1)
        ws3.conditional_format(f"A2:{last3}{nrow3+1}",{"type":"formula","criteria":'=INDIRECT("$A"&ROW())="New"',"format":red})
        ws3.freeze_panes(1,1); ws3.autofilter(0,1,nrow3,ncol3-1); autofit(ws3,crd_df,skip_cols=[0])

        # CRDPD_Mth_Sizes
        crdpd_df.to_excel(writer,"CRDPD_Mth_Sizes",index=False)
        ws4=writer.sheets["CRDPD_Mth_Sizes"]; nrow4,ncol4=crdpd_df.shape
        ws4.set_row(0,None,header); ws4.set_column(0,0,0)
        last4=excel_col(ncol4-1)
        ws4.conditional_format(f"A2:{last4}{nrow4+1}",{"type":"formula","criteria":'=INDIRECT("$A"&ROW())="New"',"format":red})
        ws4.freeze_panes(1,1); ws4.autofilter(0,1,nrow4,ncol4-1); autofit(ws4,crdpd_df,skip_cols=[0])

    return output.getvalue()

excel_bytes = build_excel_bytes(df_sizelist,sizes_df,crd_df,crdpd_df)
st.download_button("⬇️ Download Excel (match Colab)", data=excel_bytes,
                   file_name="df_sizelist_ready.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
