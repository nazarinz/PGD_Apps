import re
import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO

# -------------------- App Config --------------------
st.set_page_config(page_title="PGD Apps — Tooling Sizelist Generator", page_icon="📊", layout="wide")
st.title("📊 Subtotal Generator — Sizes, CRD_Mth, CRDPD_Mth")

# -------------------- Sidebar --------------------
with st.sidebar:
    st.markdown("### ⚙️ Pengaturan")
    st.markdown("- Upload file template sesuai kolom wajib.")
    st.markdown("- Pilih tanggal **New Order**.")
    st.markdown("---")

# -------------------- Helpers --------------------
@st.cache_data(show_spinner=False)
def load_excel(file) -> pd.DataFrame:
    return pd.read_excel(file)

def insert_after(df: pd.DataFrame, after_col: str, new_col: str, values):
    pos = df.columns.get_loc(after_col) + 1
    df.insert(pos, new_col, values)

def day_to_class(day_series: pd.Series, suffix: str) -> pd.Series:
    out = np.where(day_series >= 24, f"30_{suffix}",
          np.where(day_series >= 16, f"23_{suffix}",
          np.where(day_series >= 8,  f"15_{suffix}",
          np.where(day_series >= 1,  f"07_{suffix}", None))))
    return pd.Series(out, index=day_series.index, dtype="object")

def fmt_mdy(dt) -> str:
    dt = pd.to_datetime(dt)
    return f"{dt.month}/{dt.day}/{dt.year}"

def make_subtotal_only(df_source, group_col, order_cols, label_fmt):
    """
    Build dataframe subtotal-only per group, tanpa detail, + Grand Total.
    Kolom pertama = Remark (untuk conditional font merah saat export).
    """
    if group_col not in df_source.columns:
        raise ValueError(f"Kolom grup '{group_col}' tidak ada di data sumber.")

    # Sort by group (na last)
    df_vis = df_source[[group_col] + order_cols].copy().sort_values(group_col, ascending=True, na_position="last")

    # Sisip Remark utk flag merah (index sudah dari df_source terfilter)
    df_work = pd.concat(
        [df_source.loc[df_vis.index, ["Remark"]].reset_index(drop=True),
         df_vis.reset_index(drop=True)],
        axis=1
    )

    pieces = []
    for key, grp in df_work.groupby(df_work[group_col], dropna=False, sort=True):
        subtotal = {col: "" for col in df_work.columns}
        label = "(blank)" if pd.isna(key) else label_fmt(key)
        has_new = (grp["Remark"].astype(str).str.lower() == "new").any()
        subtotal["Remark"] = "New" if has_new else "cfm"
        subtotal[group_col] = label
        for col in order_cols:
            subtotal[col] = grp[col].sum(skipna=True)
        pieces.append(pd.DataFrame([subtotal], columns=df_work.columns))

    out = pd.concat(pieces, ignore_index=True)

    # Grand total
    grand_vals = df_vis.drop(columns=[group_col]).sum(numeric_only=True)
    grand = {col:"" for col in out.columns}
    grand[group_col] = "Grand Total"
    grand["Remark"]  = "cfm"
    for col in order_cols:
        grand[col] = grand_vals.get(col, "")
    out = pd.concat([out, pd.DataFrame([grand])], ignore_index=True)
    return out

def build_excel_bytes(df_sizelist, sizes_df, crd_df, crdpd_df) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="m/d/yyyy") as writer:
        wb = writer.book

        # Styles
        base   = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter"})
        datef  = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","num_format":"m/d/yyyy"})
        header = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","bold":True})
        red    = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","font_color":"#FF0000"})
        num0   = wb.add_format({"font_name":"Calibri","font_size":9,"align":"center","valign":"vcenter","num_format":"0"})

        def excel_col(i):
            s=""; n=i+1
            while n: n, r = divmod(n-1, 26); s = chr(65+r)+s
            return s

        def autofit(ws, df, skip_cols=[]):
            for j, col in enumerate(df.columns):
                if j in skip_cols: 
                    continue
                maxlen = max(len(str(col)), df[col].astype(str).map(len).max() if len(df)>0 else 0)
                ws.set_column(j, j, max(8, min(60, maxlen + 2)))

        # --- SHEET Data ---
        df_sizelist.to_excel(writer, sheet_name="Data", index=False)
        ws1 = writer.sheets["Data"]
        nrow1, ncol1 = df_sizelist.shape
        ws1.set_column(0, ncol1-1, None, base); ws1.set_row(0, None, header)

        # Kolom tanggal → short date
        dt_cols = [c for c in df_sizelist.columns if np.issubdtype(df_sizelist[c].dtype, np.datetime64)]
        for c in dt_cols:
            idx = df_sizelist.columns.get_loc(c)
            ws1.set_column(idx, idx, 12, datef)

        # Kolom angka → 0 desimal (Order Quantity + UK_*)
        num_cols = []
        if "Order Quantity" in df_sizelist.columns:
            num_cols.append(df_sizelist.columns.get_loc("Order Quantity"))
        for j,c in enumerate(df_sizelist.columns):
            if re.match(r'(?i)^UK_', str(c)):
                num_cols.append(j)
        for idx in sorted(set(num_cols)):
            ws1.set_column(idx, idx, None, num0)

        # Font merah jika Remark="New"
        if "Remark" in df_sizelist.columns:
            last = excel_col(ncol1-1)
            rng  = f"A2:{last}{nrow1+1}"
            ridx = df_sizelist.columns.get_loc("Remark")
            rcol = excel_col(ridx)
            ws1.conditional_format(rng, {"type":"formula","criteria":f'=INDIRECT("${rcol}" & ROW())="New"', "format": red})

        ws1.freeze_panes(1,0); ws1.autofilter(0,0,nrow1,ncol1-1)
        autofit(ws1, df_sizelist)

        # --- SHEET Sizes (Document Date subtotals only) ---
        sizes_df.to_excel(writer, sheet_name="Sizes", index=False)
        ws2 = writer.sheets["Sizes"]
        nrow2, ncol2 = sizes_df.shape
        ws2.set_row(0, None, header)
        ws2.set_column(0, 0, 0)  # hide Remark
        last2 = excel_col(ncol2-1)
        ws2.conditional_format(f"A2:{last2}{nrow2+1}", {"type":"formula","criteria":'=INDIRECT("$A"&ROW())="New"', "format": red})
        ws2.freeze_panes(1,1); ws2.autofilter(0,1,nrow2,ncol2-1)
        autofit(ws2, sizes_df, skip_cols=[0])

        # --- SHEET CRD_Mth_Sizes (subtotals only) ---
        crd_df.to_excel(writer, sheet_name="CRD_Mth_Sizes", index=False)
        ws3 = writer.sheets["CRD_Mth_Sizes"]
        nrow3, ncol3 = crd_df.shape
        ws3.set_row(0, None, header)
        ws3.set_column(0, 0, 0)  # hide Remark
        last3 = excel_col(ncol3-1)
        ws3.conditional_format(f"A2:{last3}{nrow3+1}", {"type":"formula","criteria":'=INDIRECT("$A"&ROW())="New"', "format": red})
        ws3.freeze_panes(1,1); ws3.autofilter(0,1,nrow3,ncol3-1)
        autofit(ws3, crd_df, skip_cols=[0])

        # --- SHEET CRDPD_Mth_Sizes (subtotals only) ---
        crdpd_df.to_excel(writer, sheet_name="CRDPD_Mth_Sizes", index=False)
        ws4 = writer.sheets["CRDPD_Mth_Sizes"]
        nrow4, ncol4 = crdpd_df.shape
        ws4.set_row(0, None, header)
        ws4.set_column(0, 0, 0)  # hide Remark
        last4 = excel_col(ncol4-1)
        ws4.conditional_format(f"A2:{last4}{nrow4+1}", {"type":"formula","criteria":'=INDIRECT("$A"&ROW())="New"', "format": red})
        ws4.freeze_panes(1,1); ws4.autofilter(0,1,nrow4,ncol4-1)
        autofit(ws4, crdpd_df, skip_cols=[0])

    return output.getvalue()

def make_template_example() -> bytes:
    """Generate contoh template .xlsx biar user lain gampang isi."""
    cols = [
        "Document Date","LPD","CRD","PD","Order Quantity","Article",
        "UK_1","UK_1-","UK_2","UK_2-","UK_3K","UK_4K"
    ]
    df = pd.DataFrame({
        "Document Date": pd.to_datetime(["2025-09-12","2025-09-20","2025-09-20","2025-09-25"]),
        "LPD":          [pd.NaT, pd.to_datetime("2025-09-18"), pd.NaT, pd.to_datetime("2025-09-24")],
        "CRD":          pd.to_datetime(["2025-09-10","2025-09-22","2025-09-19","2025-09-22"]),
        "PD":           pd.to_datetime(["2025-09-11","2025-09-21","2025-09-17","2025-09-20"]),
        "Order Quantity":[100, 200, 150, 120],
        "Article":      ["HS0100000013399","FG0100000021274","HL0100000011352","HU0100000017749"],
        "UK_1":[10,20,15,12],"UK_1-":[5,10,8,7],"UK_2":[12,18,9,10],"UK_2-":[7,11,6,5],
        "UK_3K":[0,1,2,0],"UK_4K":[1,0,0,1]
    })[cols]
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as w:
        df.to_excel(w, index=False, sheet_name="Template")
    return bio.getvalue()

# -------------------- Upload & Input --------------------
uploaded = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])
if uploaded is None:
    c1, c2 = st.columns(2)
    with c1:
        st.info("⬆️ Silakan upload file Excel terlebih dahulu")
    with c2:
        st.download_button("📥 Download contoh template", data=make_template_example(),
                           file_name="template_pgd.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.stop()

df_sizelist = load_excel(uploaded)

# Validasi kolom wajib (termasuk 'Article' untuk filter)
required_cols = {"Document Date","LPD","CRD","PD","Order Quantity","Article"}
missing = sorted(col for col in required_cols if col not in df_sizelist.columns)
if missing:
    st.error(f"Kolom wajib hilang: {', '.join(missing)}. Silakan gunakan template yang benar.")
    st.stop()

# ===== NEW: Filter Article -> gunakan HS & FG saja (drop HL/HU & lainnya) =====
mask_keep = df_sizelist["Article"].astype(str).str.startswith(("HS","FG"))
df_sizelist = df_sizelist[mask_keep].copy()
df_sizelist.reset_index(drop=True, inplace=True)
st.caption(f"✅ Filter Article diterapkan: {mask_keep.sum()} baris dipertahankan (HS/FG).")

# Input tanggal New Order (sidebar)
with st.sidebar:
    new_order_date = st.date_input("Tanggal New Order terakhir", pd.to_datetime("today"))
NEW_ORDER_DATE = pd.to_datetime(new_order_date)

# -------------------- Normalisasi tanggal penting --------------------
for col in ["Document Date", "LPD", "CRD", "PD"]:
    df_sizelist[col] = pd.to_datetime(df_sizelist[col], errors="coerce")

# -------------------- Remark (sesuai brief) --------------------
if "Remark" in df_sizelist.columns:
    df_sizelist.drop(columns=["Remark"], inplace=True)

remark = np.where(
    (df_sizelist["Document Date"] >= NEW_ORDER_DATE) & (df_sizelist["LPD"].isna()),
    "New", "cfm"
)
ins_pos = df_sizelist.columns.get_loc("Document Date") + 1
df_sizelist.insert(ins_pos, "Remark", remark)

# -------------------- LPD kosong = min(CRD, PD) --------------------
row_min = pd.concat([df_sizelist["CRD"], df_sizelist["PD"]], axis=1).min(axis=1, skipna=True)
df_sizelist.loc[df_sizelist["LPD"].isna(), "LPD"] = row_min[df_sizelist["LPD"].isna()]

# -------------------- Turunan CRD (ikut Remark) --------------------
YM_CRD  = df_sizelist["CRD"].dt.strftime("%Y%m")
Day_CRD = df_sizelist["CRD"].dt.day
is_new  = df_sizelist["Remark"].astype(str).str.lower().eq("new")
Class_CRD = np.where(is_new, day_to_class(Day_CRD,"new"), day_to_class(Day_CRD,"cfm"))
CRD_Mth = YM_CRD.fillna("").astype(str) + pd.Series(Class_CRD).fillna("")

insert_after(df_sizelist, "CRD", "YM_CRD", YM_CRD)
insert_after(df_sizelist, "YM_CRD", "Day_CRD", Day_CRD)
insert_after(df_sizelist, "Day_CRD", "Class_CRD", Class_CRD)
insert_after(df_sizelist, "Class_CRD", "CRD_Mth", CRD_Mth)

# -------------------- Turunan LPD (ikut Remark) --------------------
YM_CRDPD  = df_sizelist["LPD"].dt.strftime("%Y%m")
Day_CRDPD = df_sizelist["LPD"].dt.day
suffix    = np.where(df_sizelist["Remark"].str.lower()=="new", "new", "cfm")
Class_CRDPD = np.where(suffix=="new", day_to_class(Day_CRDPD,"new"), day_to_class(Day_CRDPD,"cfm"))
CRDPD_Mth = YM_CRDPD.fillna("").astype(str) + pd.Series(Class_CRDPD).fillna("")

insert_after(df_sizelist, "LPD", "YM_CRDPD", YM_CRDPD)
insert_after(df_sizelist, "YM_CRDPD", "Day_CRDPD", Day_CRDPD)
insert_after(df_sizelist, "Day_CRDPD", "Class_CRDPD", Class_CRDPD)
insert_after(df_sizelist, "Class_CRDPD", "CRDPD_Mth", CRDPD_Mth)

# -------------------- Build 3 sheet subtotal-only --------------------
size_cols = [c for c in df_sizelist.columns if re.match(r'(?i)^UK_', str(c))]
order_cols = ["Order Quantity"] + size_cols

sizes_df  = make_subtotal_only(df_sizelist, "Document Date", order_cols, label_fmt=lambda k: f"{fmt_mdy(k)} Total")
crd_df    = make_subtotal_only(df_sizelist, "CRD_Mth",      order_cols, label_fmt=lambda k: str(k))
crdpd_df  = make_subtotal_only(df_sizelist, "CRDPD_Mth",    order_cols, label_fmt=lambda k: str(k))
crdpd_df  = crdpd_df.rename(columns={"CRDPD_Mth": "CRDPD_month"})

# -------------------- Preview --------------------
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

# -------------------- Unduh Excel (format match Colab) --------------------
excel_bytes = build_excel_bytes(df_sizelist, sizes_df, crd_df, crdpd_df)
st.download_button("⬇️ Download Excel (match Colab)",
                   data=excel_bytes,
                   file_name="df_sizelist_ready.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
