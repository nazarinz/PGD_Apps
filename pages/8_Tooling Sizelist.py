import re
import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="PGD Apps — Tooling Sizelist", page_icon="📊", layout="wide")
st.title("📊 PGD Apps — Tooling Sizelist")

# === Upload File ===
uploaded = st.file_uploader("Upload file Excel", type=["xlsx", "xls"])
if uploaded is None:
    st.info("⬆️ Silakan upload file Excel terlebih dahulu")
    st.stop()

# === Load Data ===
df_sizelist = pd.read_excel(uploaded)

# === Input tanggal New Order ===
new_order_date = st.date_input("Tanggal New Order terakhir", pd.to_datetime("today"))
NEW_ORDER_DATE = pd.to_datetime(new_order_date)

# === Pastikan datetime ===
for col in ["Document Date", "LPD", "CRD", "PD"]:
    if col in df_sizelist.columns:
        df_sizelist[col] = pd.to_datetime(df_sizelist[col], errors="coerce")

# === Remark ===
if "Remark" in df_sizelist.columns:
    df_sizelist.drop(columns=["Remark"], inplace=True)

remark = np.where(
    (df_sizelist["Document Date"] >= NEW_ORDER_DATE) & (df_sizelist["LPD"].isna()),
    "New", "cfm"
)
pos_remark = df_sizelist.columns.get_loc("Document Date") + 1
df_sizelist.insert(pos_remark, "Remark", remark)

# === Isi LPD kosong ===
if {"CRD","PD","LPD"}.issubset(df_sizelist.columns):
    row_min = pd.concat([df_sizelist["CRD"], df_sizelist["PD"]], axis=1).min(axis=1, skipna=True)
    mask_na = df_sizelist["LPD"].isna()
    df_sizelist.loc[mask_na, "LPD"] = row_min[mask_na]

# === Helper fungsi ===
def insert_after(df, after_col, new_col, values):
    pos = df.columns.get_loc(after_col) + 1
    df.insert(pos, new_col, values)

def day_to_class(day_series, suffix):
    return np.where(day_series >= 24, f"30_{suffix}",
           np.where(day_series >= 16, f"23_{suffix}",
           np.where(day_series >= 8,  f"15_{suffix}",
           np.where(day_series >= 1,  f"07_{suffix}", None))))

# === Turunan CRD ===
if "CRD" in df_sizelist.columns:
    YM_CRD  = df_sizelist["CRD"].dt.strftime("%Y%m")
    Day_CRD = df_sizelist["CRD"].dt.day
    is_new  = df_sizelist["Remark"].astype(str).str.lower().eq("new")
    Class_CRD = np.where(is_new, day_to_class(Day_CRD,"new"), day_to_class(Day_CRD,"cfm"))
    CRD_Mth = YM_CRD.fillna("").astype(str) + pd.Series(Class_CRD).fillna("")
    insert_after(df_sizelist,"CRD","YM_CRD",YM_CRD)
    insert_after(df_sizelist,"YM_CRD","Day_CRD",Day_CRD)
    insert_after(df_sizelist,"Day_CRD","Class_CRD",Class_CRD)
    insert_after(df_sizelist,"Class_CRD","CRD_Mth",CRD_Mth)

# === Turunan LPD ===
if "LPD" in df_sizelist.columns:
    YM_CRDPD  = df_sizelist["LPD"].dt.strftime("%Y%m")
    Day_CRDPD = df_sizelist["LPD"].dt.day
    suffix    = np.where(df_sizelist["Remark"].str.lower()=="new","new","cfm")
    Class_CRDPD = np.where(suffix=="new", day_to_class(Day_CRDPD,"new"), day_to_class(Day_CRDPD,"cfm"))
    CRDPD_Mth = YM_CRDPD.fillna("").astype(str) + pd.Series(Class_CRDPD).fillna("")
    insert_after(df_sizelist,"LPD","YM_CRDPD",YM_CRDPD)
    insert_after(df_sizelist,"YM_CRDPD","Day_CRDPD",Day_CRDPD)
    insert_after(df_sizelist,"Day_CRDPD","Class_CRDPD",Class_CRDPD)
    insert_after(df_sizelist,"Class_CRDPD","CRDPD_Mth",CRDPD_Mth)

# === Utility subtotal-only ===
def make_subtotal_only(df_source, group_col_name, order_cols, label_fmt):
    df_vis = df_source[[group_col_name]+order_cols].copy()
    df_vis = df_vis.sort_values(by=group_col_name, ascending=True, na_position="last")
    df_flag = pd.DataFrame({"Remark": df_source["Remark"]})
    df_work = pd.concat([df_flag.loc[df_vis.index].reset_index(drop=True),
                         df_vis.reset_index(drop=True)], axis=1)
    pieces=[]
    for key,grp in df_work.groupby(df_work[group_col_name], dropna=False, sort=True):
        subtotal={col:"" for col in df_work.columns}
        label = "Subtotal (blank)" if pd.isna(key) else label_fmt(key)
        has_new = (grp["Remark"].str.lower()=="new").any()
        subtotal["Remark"]="New" if has_new else "cfm"
        subtotal[group_col_name]=label
        for col in order_cols:
            subtotal[col]=grp[col].sum(skipna=True)
        pieces.append(pd.DataFrame([subtotal],columns=df_work.columns))
    df_out=pd.concat(pieces,ignore_index=True)
    grand_vals=df_vis.drop(columns=[group_col_name]).sum(numeric_only=True)
    grand={col:"" for col in df_out.columns}
    grand[group_col_name]="Grand Total"; grand["Remark"]="cfm"
    for col in order_cols: grand[col]=grand_vals.get(col,"")
    df_out=pd.concat([df_out,pd.DataFrame([grand])],ignore_index=True)
    return df_out

# === Build 3 subtotal sheets ===
size_cols = [c for c in df_sizelist.columns if re.match(r'(?i)^UK_', str(c))]
order_cols=["Order Quantity"]+size_cols

sizes_df = make_subtotal_only(df_sizelist,"Document Date",order_cols,
    label_fmt=lambda k:f"{pd.to_datetime(k).strftime('%-m/%-d/%Y')} Total")
crd_df   = make_subtotal_only(df_sizelist,"CRD_Mth",order_cols,label_fmt=lambda k:str(k))
crdpd_df = make_subtotal_only(df_sizelist,"CRDPD_Mth",order_cols,label_fmt=lambda k:str(k))
crdpd_df=crdpd_df.rename(columns={"CRDPD_Mth":"CRDPD_month"})

# === Preview di Streamlit ===
st.subheader("📑 Data Preview")
st.dataframe(df_sizelist.head(20))

st.subheader("📑 Sizes (Subtotal Only)")
st.dataframe(sizes_df.head(20))

st.subheader("📑 CRD_Mth_Sizes (Subtotal Only)")
st.dataframe(crd_df.head(20))

st.subheader("📑 CRDPD_Mth_Sizes (Subtotal Only)")
st.dataframe(crdpd_df.head(20))

# === Export to Excel ===
output=BytesIO()
with pd.ExcelWriter(output,engine="xlsxwriter") as writer:
    df_sizelist.to_excel(writer,sheet_name="Data",index=False)
    sizes_df.to_excel(writer,sheet_name="Sizes",index=False)
    crd_df.to_excel(writer,sheet_name="CRD_Mth_Sizes",index=False)
    crdpd_df.to_excel(writer,sheet_name="CRDPD_Mth_Sizes",index=False)
st.download_button("⬇️ Download Excel", data=output.getvalue(),
                   file_name="df_sizelist_ready.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
