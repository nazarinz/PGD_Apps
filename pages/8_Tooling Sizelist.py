from utils.auth import require_login

require_login()

# ==========================================
# 8_Tooling Sizelist.py — PGD Apps (v12)
# ==========================================
import re
import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO

# ================= Config =================
REQUIRED_COLS = ["Sales Order", "Document Date", "Article", "Order Quantity",
                 "CRD", "PD", "LPD", "Working Status"]
ARTICLE_PREFIXES = ("FG", "HS")
WS_NEW_VALUE = 10
DATE_COLS = ["Document Date", "LPD", "CRD", "PD"]

# ==========================================
st.set_page_config(page_title="PGD Apps — Tooling Sizelist", page_icon="📊", layout="wide")
st.title("📊 PGD Tooling Sizelist — Subtotal Generator (v12)")

# ================= Upload & Input =================
uploaded = st.file_uploader("📤 Upload file Excel (SAP/In-house Sizelist)", type=["xlsx", "xls"])
if uploaded is None:
    st.info("⬆️ Silakan upload file Excel terlebih dahulu sebelum melanjutkan.")
    st.stop()

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
| **Working Status**  | Status pengerjaan (wajib untuk deteksi "New")               |
| **UK_***            | Kolom ukuran (size breakdown)                               |
""")

# ================= Baca & Validasi =================
df_raw = pd.read_excel(uploaded)

missing_cols = [c for c in REQUIRED_COLS if c not in df_raw.columns]
if missing_cols:
    st.error(f"❌ Kolom berikut tidak ditemukan di file: **{', '.join(missing_cols)}**")
    st.stop()

with st.expander("🔍 DEBUG 1: Info Data Mentah (sebelum filter)", expanded=False):
    st.write(f"**Total rows awal:** {len(df_raw)}")
    st.write(f"**Kolom:** {list(df_raw.columns)}")
    st.write("**Sample Article:**", df_raw["Article"].head(10).tolist())
    st.write("**Working Status (value counts):**")
    st.write(df_raw["Working Status"].value_counts())
    st.write("**Document Date sample:**")
    st.write(df_raw["Document Date"].head(10))

# Filter Article FG / HS
before = len(df_raw)
df = df_raw[df_raw["Article"].astype(str).str.startswith(ARTICLE_PREFIXES)].copy()
st.info(f"✅ Filter Article {'/'.join(ARTICLE_PREFIXES)}: {before} rows → {len(df)} rows")

if df.empty:
    st.error("❌ Tidak ada data setelah filter Article. Periksa kolom Article di file kamu.")
    st.stop()

# ================= Input Eksekusi =================
st.subheader("⚙️ Pengaturan Eksekusi")
new_order_date = st.date_input("Tanggal New Order terakhir *wajib diisi*", value=None)

cancel_sos_input = st.text_area(
    "Daftar Sales Order Cancel (pisahkan dengan koma ATAU baris baru):",
    placeholder="contoh:\n11184283\n11185888\n11327194\natau: 11184283, 11185888, 11327194"
)

if not st.button("🚀 Execute Generate"):
    st.stop()

if not new_order_date:
    st.error("❌ Silakan isi tanggal New Order terlebih dahulu.")
    st.stop()

NEW_ORDER_DATE = pd.to_datetime(new_order_date)

# ✅ FIX: Split koma DAN newline — support paste multi-baris
cancel_sos: list[str] = [s.strip() for s in re.split(r'[,\n\r]+', cancel_sos_input) if s.strip()]

def normalize_so(val) -> str:
    """Normalize SO ke string bersih — handle float artifact '11184283.0' → '11184283'."""
    s = str(val).strip()
    return s[:-2] if s.endswith(".0") else s

st.success(f"✅ NEW_ORDER_DATE: **{NEW_ORDER_DATE.strftime('%Y-%m-%d')}**")
if cancel_sos:
    st.warning(f"⚠️ Cancel SO yang akan ditandai ({len(cancel_sos)} SO): {', '.join(cancel_sos)}")
else:
    st.info("ℹ️ Tidak ada SO Cancel yang dimasukkan.")

# ================= Normalisasi Tanggal =================
for col in DATE_COLS:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors="coerce")

# ================= DEBUG 2 =================
with st.expander("🔍 DEBUG 2: Kondisi Sebelum Remark Dibuat", expanded=True):
    st.write(f"**NEW_ORDER_DATE:** {NEW_ORDER_DATE}")

    ws_numeric_debug = pd.to_numeric(df["Working Status"], errors="coerce").fillna(-1)
    lpd_null = df["LPD"].isna()
    cond1 = df["Document Date"] >= NEW_ORDER_DATE
    cond2 = lpd_null
    cond3 = ws_numeric_debug == WS_NEW_VALUE

    st.write("### 🎯 Ringkasan Kondisi:")
    st.write(f"- Cond 1 (Doc Date >= NEW): **{cond1.sum()}** rows")
    st.write(f"- Cond 2 (LPD null): **{cond2.sum()}** rows")
    st.write(f"- Cond 3 (WS == {WS_NEW_VALUE}): **{cond3.sum()}** rows")
    st.write(f"- **Akan jadi 'New': {(cond1 & cond2 & cond3).sum()} rows**")

    debug_df = pd.DataFrame({
        "Sales Order": df["Sales Order"],
        "Doc Date": df["Document Date"],
        "Doc >= NEW?": cond1,
        "LPD": df["LPD"],
        "LPD null?": cond2,
        "WS (raw)": df["Working Status"],
        "WS numeric": ws_numeric_debug,
        f"WS == {WS_NEW_VALUE}?": cond3,
    })
    st.dataframe(debug_df.head(20), use_container_width=True)

# ================= Remark =================
if "Remark" in df.columns:
    df.drop(columns=["Remark"], inplace=True)

ws_numeric = pd.to_numeric(df["Working Status"], errors="coerce").fillna(-1)
remark_values = np.where(
    (df["Document Date"] >= NEW_ORDER_DATE)
    & (df["LPD"].isna())
    & (ws_numeric == WS_NEW_VALUE),
    "New", "cfm"
)
ins_pos = df.columns.get_loc("Document Date") + 1
df.insert(ins_pos, "Remark", remark_values)

with st.expander("🔍 DEBUG 3: Hasil Remark", expanded=True):
    st.write("**Distribusi Remark:**")
    st.write(pd.Series(remark_values).value_counts())
    st.dataframe(df[["Sales Order", "Document Date", "LPD", "Working Status", "Remark"]].head(20))

# ================= Isi LPD Kosong =================
if {"CRD", "PD", "LPD"}.issubset(df.columns):
    mask_lpd_null = df["LPD"].isna()
    row_min = pd.concat([df["CRD"], df["PD"]], axis=1).min(axis=1, skipna=True)
    df.loc[mask_lpd_null, "LPD"] = row_min[mask_lpd_null]
    st.info(f"✅ LPD kosong diisi dengan min(CRD, PD): {mask_lpd_null.sum()} rows")

# ================= Helpers =================
def insert_after(df: pd.DataFrame, after_col: str, new_col: str, values) -> None:
    pos = df.columns.get_loc(after_col) + 1
    df.insert(pos, new_col, values)

def bucket_day(day_series: pd.Series) -> np.ndarray:
    day = day_series.values
    return np.where(day >= 24, "30",
           np.where(day >= 16, "23",
           np.where(day >= 8,  "15",
           np.where(day >= 1,  "07", None))))

# ================= CRD_Mth & CRDPD_Mth =================
if "CRD" in df.columns:
    YM_CRD   = df["CRD"].dt.strftime("%Y%m")
    Day_CRD  = df["CRD"].dt.day
    Cls_CRD  = pd.Series(bucket_day(Day_CRD), index=df.index)
    Rem      = df["Remark"].str.lower().fillna("cfm")
    CRD_Mth  = YM_CRD.fillna("") + Cls_CRD.fillna("") + "_" + Rem
    insert_after(df, "CRD",       "YM_CRD",    YM_CRD)
    insert_after(df, "YM_CRD",    "Day_CRD",   Day_CRD)
    insert_after(df, "Day_CRD",   "Class_CRD", Cls_CRD)
    insert_after(df, "Class_CRD", "CRD_Mth",   CRD_Mth)

if "LPD" in df.columns:
    YM_LPD    = df["LPD"].dt.strftime("%Y%m")
    Day_LPD   = df["LPD"].dt.day
    Cls_LPD   = pd.Series(bucket_day(Day_LPD), index=df.index)
    Rem       = df["Remark"].str.lower().fillna("cfm")
    CRDPD_Mth = YM_LPD.fillna("") + Cls_LPD.fillna("") + "_" + Rem
    insert_after(df, "LPD",        "YM_CRDPD",   YM_LPD)
    insert_after(df, "YM_CRDPD",   "Day_CRDPD",  Day_LPD)
    insert_after(df, "Day_CRDPD",  "Class_CRDPD", Cls_LPD)
    insert_after(df, "Class_CRDPD","CRDPD_Mth",  CRDPD_Mth)

# ================= Subtotal Builder =================
size_cols  = [c for c in df.columns if re.match(r'(?i)^UK_', str(c))]
order_cols = ["Order Quantity"] + size_cols
cancel_set = set(cancel_sos)  # user input sudah string bersih

# Normalisasi kolom SO di dataframe agar cocok saat compare
df["_SO_str"] = df["Sales Order"].apply(normalize_so)

def make_subtotal(df: pd.DataFrame, group_col: str) -> tuple[pd.DataFrame, list[str]]:
    pieces, color_tags = [], []
    for key, grp in df.groupby(group_col, dropna=False):
        has_new    = (grp["Remark"].str.lower() == "new").any()
        has_cancel = grp["_SO_str"].isin(cancel_set).any()
        color      = "red" if has_new else ("purple" if has_cancel else "black")

        row = {group_col: key}
        row["Remark"] = "New" if has_new else "cfm"
        for col in order_cols:
            row[col] = grp[col].sum(skipna=True)

        pieces.append(row)
        color_tags.append(color)

    out = pd.DataFrame(pieces).drop(columns=["Remark"], errors="ignore")
    return out, color_tags

sizes_df,  color_sizes  = make_subtotal(df, "Document Date")
crd_df,    color_crd    = make_subtotal(df, "CRD_Mth")
crdpd_df,  color_crdpd  = make_subtotal(df, "CRDPD_Mth")

# ================= Warna Per Row Data Utama =================
def row_colors(df: pd.DataFrame, cancel_set: set[str]) -> list[str]:
    colors = []
    for _, row in df.iterrows():
        so = str(row.get("_SO_str", row.get("Sales Order", "")))
        if so in cancel_set:
            colors.append("purple")
        elif str(row.get("Remark", "")).lower() == "new":
            colors.append("red")
        else:
            colors.append("black")
    return colors

color_main = row_colors(df, cancel_set)

# ================= Preview =================
st.subheader("📑 Data (preview 20 rows)")
st.dataframe(df.head(20), use_container_width=True)

# ================= Excel Export =================
def build_excel() -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="m/d/yyyy") as writer:
        wb = writer.book

        def fmt_text(color: str = "black"):
            return wb.add_format({"font_name": "Calibri", "font_size": 9,
                                  "align": "center", "valign": "vcenter",
                                  "font_color": color})

        def fmt_date(color: str = "black"):
            return wb.add_format({"font_name": "Calibri", "font_size": 9,
                                  "align": "center", "valign": "vcenter",
                                  "font_color": color, "num_format": "m/d/yyyy"})

        def fmt_num(color: str = "black"):
            return wb.add_format({"font_name": "Calibri", "font_size": 9,
                                  "align": "center", "valign": "vcenter",
                                  "font_color": color, "num_format": "0"})

        def write_sheet(ws, data: pd.DataFrame, colors: list[str]) -> None:
            for j, col in enumerate(data.columns):
                ws.write(0, j, col, fmt_text("black"))
            for i, (_, row) in enumerate(data.iterrows()):
                c = colors[i]
                for j, val in enumerate(row):
                    if pd.isna(val):
                        ws.write(i + 1, j, "", fmt_text(c))
                    elif isinstance(val, (pd.Timestamp, np.datetime64)):
                        ws.write_datetime(i + 1, j, pd.to_datetime(val), fmt_date(c))
                    elif isinstance(val, (int, float, np.number)) and not isinstance(val, bool):
                        ws.write_number(i + 1, j, float(val), fmt_num(c))
                    else:
                        ws.write(i + 1, j, str(val), fmt_text(c))
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, len(data), len(data.columns) - 1)

        sheets = [
            ("Data",           df.drop(columns=["_SO_str"], errors="ignore"),         color_main),
            ("Sizes",          sizes_df,   color_sizes),
            ("CRD_Mth_Sizes",  crd_df,     color_crd),
            ("CRDPD_Mth_Sizes",crdpd_df,   color_crdpd),
        ]
        for sheet_name, data, colors in sheets:
            data.to_excel(writer, sheet_name=sheet_name, index=False)
            write_sheet(writer.sheets[sheet_name], data, colors)

        # Summary
        total_so     = df["Sales Order"].nunique() if "Sales Order" in df else 0
        total_new    = (df["Remark"].str.lower() == "new").sum()
        total_cancel = df["_SO_str"].isin(cancel_set).sum()
        summary = pd.DataFrame({
            "Metric": ["Total SO", "Total New", "Total Cancel"],
            "Value":  [total_so, int(total_new), int(total_cancel)]
        })
        summary.to_excel(writer, sheet_name="Summary", index=False)
        ws_sum = writer.sheets["Summary"]
        ws_sum.set_column(0, 0, 18)
        ws_sum.set_column(1, 1, 12)

    return output.getvalue()

excel_bytes = build_excel()

st.download_button(
    "⬇️ Download Excel",
    data=excel_bytes,
    file_name="Tooling_Sizelist_v12.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.success("✅ Processing selesai! Cancel SO per baris / per koma keduanya sudah support.")
