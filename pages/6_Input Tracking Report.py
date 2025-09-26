# ============================
# PGD Apps — Input Tracking Report (Reroute + Pending Cancel + No-Pivot Formatter)
# ============================

import re
import os
import pandas as pd
import streamlit as st
from utils import set_page, header, footer, write_excel_autofit

# --- Urutan kolom yang diinginkan (Reroute) ---
DESIRED_ORDER = [
    "Working Status","Select","Working Status Descr.","Requirement Segment",
    "Sales Order","Sold-To PO No.","CRD","CRD-DRC","PD","POSDD-DRC","POSDD",
    "FPD","FPD-DRC","PODD","PODD-DRC","Est. Inspection Date",
    "LPD","LPD-DRC","FGR","Cust Article No.","Model Name","Article",
    "Lead Time","Season","Product Hierarchy 3","Ship-To Search Term",
    "Ship-To Country","Document Date","Order Quantity","Order Type"
]

# --- Kolom target & urutan minimal (Pending Cancel — Filter & Reorder) ---
TARGET_ORDER = [
    "Sales Order","Remark","Sold-To PO No.","Ship-To Party PO No.",
    "Model Name","Cust Article No.","Article","Order Quantity",
    "CRD","PD","LPD","PODD","Ship-To Search Term","Ship-To Name"
]

# --- (Baru) Urutan kolom akhir hasil NO-PIVOT (Pending Cancel Format) ---
FINAL_ORDER_NOPIVOT = [
    "Ticket Date","Work Center","Document Date","Sales Order","Customer Contract ID",
    "Sold-To PO No.","BTP Ticket","Factory E-mail Subject","Model Name","Cust Article No.",
    "Article","Ship-To Search Term","Ship-To Country","Size","Order Quantity",
    "Reduce Qty","Increase Qty","New Qty","LPD","PODD","Status","Cost Category","Claim Cost"
]

# --- (Opsional) Tabel panduan kolom INPUT yang umum dipakai untuk No-Pivot ---
NOPIVOT_INPUT_EXPECTED = [
    "Work Center","Order Type","Requirement Segment","Site","Sales Order","BTP Ticket",
    "Customer Contract ID","Sold-To PO No.","Status","Cost Category","Claim Cost",
    "Ship-To Party PO No.","Article","Cust Article No.","Article Lead Time","Model Name",
    "CRD","PD","PODD","LPD","Ship-To No.","Ship-To Search Term","Ship-To Name",
    "Ship-To Country","Document Date","Remark","Order Quantity"
]

# =========================
# Util umum (Reroute/Filter)
# =========================
def reorder_columns(df: pd.DataFrame, desired_order: list) -> tuple[pd.DataFrame, list, list]:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in desired_order if c not in df.columns]
    extra   = [c for c in df.columns if c not in desired_order]
    for col in missing:
        df[col] = pd.NA
    new_order = desired_order + [c for c in df.columns if c not in desired_order]
    df = df[new_order]
    return df, missing, extra

def filter_and_reorder(df: pd.DataFrame, target_cols: list) -> tuple[pd.DataFrame, list]:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in target_cols if c not in df.columns]
    for col in missing:
        df[col] = pd.NA
    out_df = df[target_cols]
    return out_df, missing

def _df_from_list(cols: list, title_col: str = "Kolom Wajib/Urutan"):
    return pd.DataFrame({title_col: cols}, index=range(1, len(cols) + 1))

# =========================
# Util normalisasi ringan
# =========================
def _normalize_header(name: str) -> str:
    base = re.sub(r"\s+", " ", str(name).strip())
    alias_map = {
        "BTP Ticket": "BTP Ticket",
        "Ship-To Party PO No.": "Ship-To Party PO No.",
        "Sold-To PO No.": "Sold-To PO No.",
        "Cust Article No.": "Cust Article No.",
        "Ship-To Search Term": "Ship-To Search Term",
        "Ship-To Country": "Ship-To Country",
        "Document Date": "Document Date",
        "Order Quantity": "Order Quantity",
    }
    return alias_map.get(base, base)

def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [_normalize_header(c) for c in df.columns]
    return df

def _to_numeric_series(s: pd.Series) -> pd.Series:
    """Bersihkan angka dari string (e.g. '1,200' / '$3,432.39' / '24 pcs') → numerik."""
    if not isinstance(s, pd.Series):
        s = pd.Series(s)
    cleaned = s.astype(str).str.replace(r"[^0-9\.\-]", "", regex=True).replace({"": None})
    return pd.to_numeric(cleaned, errors="coerce")

# =========================
# Formatter NO-PIVOT
# =========================
def format_no_pivot(df: pd.DataFrame, ticket_date: str, factory_subject: str) -> pd.DataFrame:
    """
    Tidak mem-pivot size. Output:
      - Size dikosongkan
      - Order Quantity diambil dari input
      - Reduce Qty = Order Quantity; Increase Qty = 0; New Qty = 0
      - Kolom ditata sesuai FINAL_ORDER_NOPIVOT
    """
    df = _normalize_columns(df).copy()

    # Pastikan kolom-kolom target ada (kalau hilang -> NA)
    for col in FINAL_ORDER_NOPIVOT:
        if col in ("Ticket Date","Factory E-mail Subject","Size","Reduce Qty","Increase Qty","New Qty"):
            continue  # akan dibuat di bawah
        if col not in df.columns:
            df[col] = pd.NA

    # Siapkan kolom turunan
    df["Ticket Date"] = ticket_date
    df["Factory E-mail Subject"] = factory_subject
    df["Size"] = ""  # sesuai permintaan, kosongkan saja

    # Order Quantity → numerik (jika tidak ada, isi 0)
    if "Order Quantity" not in df.columns:
        df["Order Quantity"] = 0
    qty_num = _to_numeric_series(df["Order Quantity"]).fillna(0)

    df["Reduce Qty"] = qty_num
    df["Increase Qty"] = 0
    df["New Qty"] = 0

    # Reorder & return
    for col in FINAL_ORDER_NOPIVOT:
        if col not in df.columns:
            df[col] = pd.NA
    return df[FINAL_ORDER_NOPIVOT]

# --- Streamlit Page ---
set_page("PGD Apps — Input Tracking Report", "📝")
header("📝 Input Tracking Report")

# ==============
# Petunjuk Umum
# ==============
with st.expander("❓ Cara pakai (singkat) — klik untuk lihat"):
    st.markdown("""
**Tujuan:** menyiapkan file Excel untuk *tracking reroute*, *tracking PO Pending Cancel (filter+reorder)*,
dan **formatter Pending Cancel (tanpa pivot size)** yang menata kolom dan nilai Qty.

**Format file:**
- Terima **.xlsx** / **.xls** (bisa multi-sheet).
- Nama kolom sebaiknya mengikuti tabel panduan di tiap sub-tools. Kolom hilang akan dibuat **NA** agar urutan tetap rapi.

**Langkah cepat:**
1) Pilih sub-tools.  
2) Upload Excel.  
3) (Khusus Formatter No-Pivot) isi **Ticket Date** & **Factory E-mail Subject**.  
4) Cek **Report/preview**, lalu **Download** hasil.
""")

choice = st.selectbox("Pilih sub-tools", [
    "Buat masukin trackingan PO Reroute",
    "Buat masukin tracking PO Pending Cancel",
    "Pending Cancel — Formatter (No Pivot)"   # <— fitur baru (tanpa pivot size)
])

# =========================
# 1) Reroute — Reorder
# =========================
if choice == "Buat masukin trackingan PO Reroute":
    st.subheader("📋 Kolom & Urutan — Trackingan PO Reroute")
    st.caption("Tool ini akan **menyusun ulang urutan kolom** sesuai daftar di bawah. Kolom yang tidak ada akan dibuat kosong (NA). Kolom tambahan (di luar daftar) dipindah ke bagian belakang.")
    st.dataframe(_df_from_list(DESIRED_ORDER, "Urutan Kolom Reroute"), use_container_width=True)

    st.markdown("---")
    up = st.file_uploader("Upload Excel (.xlsx/.xls)", type=["xlsx","xls"], accept_multiple_files=False, key="itr_reorder")
    if not up:
        st.stop()

    xls = pd.read_excel(up, sheet_name=None, engine="openpyxl")
    out_frames = {}
    report = []
    for sheet_name, df in xls.items():
        df_out, missing, extra = reorder_columns(df, DESIRED_ORDER)
        out_frames[sheet_name] = df_out
        report.append({
            "Sheet": sheet_name,
            "Missing (dibuat NA)": ", ".join(missing) if missing else "-",
            "Extra (dipindah ke belakang)": ", ".join(extra) if extra else "-",
        })
    report_df = pd.DataFrame(report)

    st.success("✅ Berhasil diproses. Cek tabel Report untuk melihat ringkasan kolom.")
    st.dataframe(report_df, use_container_width=True)

    payload = write_excel_autofit({**out_frames, "Report": report_df})
    st.download_button("⬇️ Download Hasil Reorder", data=payload,
                       file_name="reordered_output.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# =========================
# 2) Pending Cancel — Filter & Reorder Minimal
# =========================
elif choice == "Buat masukin tracking PO Pending Cancel":
    st.subheader("📋 Kolom Minimal — Tracking PO Pending Cancel")
    st.caption("Tool ini akan **filter & urutkan** hanya kolom penting di bawah ini untuk input tracking Pending Cancel. Kolom yang tidak ada akan dibuat kosong (NA).")
    st.dataframe(_df_from_list(TARGET_ORDER, "Kolom Minimal Pending Cancel"), use_container_width=True)

    st.markdown("""
**Catatan:**  
- **Sales Order, Remark, Sold-To PO No., Model Name, Article, Order Quantity, CRD, PD, LPD, PODD, Ship-To**: kolom inti untuk mencatat status *Pending Cancel*.  
- **Remark**: isi alasan/status singkat (mis. *waiting approval*, *cancel confirmed*, *reroute*, dll).  
- **Ship-To Name/Party PO No.**: wajib jika proses internal memerlukan referensi ship-to spesifik.
""")

    st.markdown("---")
    up2 = st.file_uploader("Upload Excel (.xlsx/.xls)", type=["xlsx","xls"], accept_multiple_files=False, key="itr_filter")
    if not up2:
        st.stop()

    xls = pd.read_excel(up2, sheet_name=None, engine="openpyxl")
    out_frames = {}
    report = []
    for sheet_name, df in xls.items():
        df_out, missing = filter_and_reorder(df, TARGET_ORDER)
        out_frames[sheet_name] = df_out
        report.append({
            "Sheet": sheet_name,
            "Kolom hilang (dibuat NA)": ", ".join(missing) if missing else "-",
        })
    report_df = pd.DataFrame(report)

    st.success("✅ Berhasil diproses. Cek tabel Report untuk melihat kolom yang ditambahkan (NA).")
    st.dataframe(report_df, use_container_width=True)

    payload = write_excel_autofit({**out_frames, "Report": report_df})
    st.download_button("⬇️ Download Hasil Filter+Reorder", data=payload,
                       file_name="filtered_output.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# =========================
# 3) Pending Cancel — Formatter (No Pivot)
# =========================
elif choice == "Pending Cancel — Formatter (No Pivot)":
    st.subheader("🧾 Formatter Pending Cancel (Tanpa Pivot Size)")
    st.caption("Output akan menata kolom dan nilai Qty.\n- **Size** dikosongkan\n- **Order Quantity** diambil dari input\n- **Reduce Qty = Order Quantity**, **Increase Qty = 0**, **New Qty = 0**")

    with st.expander("📥 Kolom input yang direkomendasikan (contoh)"):
        st.dataframe(_df_from_list(NOPIVOT_INPUT_EXPECTED, "Kolom Input (contoh)"), use_container_width=True)

    with st.expander("📤 Kolom output & urutan akhir"):
        st.dataframe(_df_from_list(FINAL_ORDER_NOPIVOT, "Urutan Kolom Output"), use_container_width=True)

    c1, c2, _ = st.columns([1,2,1])
    with c1:
        ticket_date = st.date_input("Ticket Date")
    with c2:
        factory_subject = st.text_input("Factory E-mail Subject", value="Pending Cancel – Summary")

    st.markdown("---")
    up3 = st.file_uploader("Upload Excel (.xlsx/.xls)", type=["xlsx","xls"], accept_multiple_files=False, key="itr_nopivot")
    if not up3:
        st.stop()

    xls = pd.read_excel(up3, sheet_name=None, engine="openpyxl")
    out_frames = {}
    report_rows = []

    for sheet_name, df in xls.items():
        final_df = format_no_pivot(
            df,
            ticket_date.strftime("%Y-%m-%d"),
            factory_subject
        )
        out_frames[sheet_name] = final_df

        total_order_qty = pd.to_numeric(final_df["Order Quantity"], errors="coerce").fillna(0).sum()
        total_reduce_qty = pd.to_numeric(final_df["Reduce Qty"], errors="coerce").fillna(0).sum()

        report_rows.append({
            "Sheet": sheet_name,
            "Rows": len(final_df),
            "Total Order Qty": total_order_qty,
            "Total Reduce Qty": total_reduce_qty
        })

    report_df = pd.DataFrame(report_rows).sort_values("Sheet")
    st.success("✅ Formatter selesai. Berikut ringkasannya:")
    st.dataframe(report_df, use_container_width=True)

    payload = write_excel_autofit({**out_frames, "Report": report_df})
    st.download_button("⬇️ Download Hasil Formatter (No Pivot)", data=payload,
                       file_name="pending_cancel_no_pivot.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

footer()
