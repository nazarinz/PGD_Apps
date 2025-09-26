# ============================
# PGD Apps — Input Tracking Report (Reroute + Pending Cancel + Pivot Size)
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

# --- (Baru) Urutan kolom akhir hasil PIVOT (Pending Cancel Format) ---
FINAL_ORDER = [
    "Ticket Date","Work Center","Document Date","Sales Order","Customer Contract ID",
    "Sold-To PO No.","BTP Ticket","Factory E-mail Subject","Model Name","Cust Article No.",
    "Article","Ship-To Search Term","Ship-To Country","Size","Qty",
    "Reduce Qty","Increase Qty","New Qty","LPD","PODD","Status","Cost Category","Claim Cost"
]

# --- (Baru) Tabel panduan kolom INPUT untuk Pivot (sesuai contohmu) ---
PIVOT_INPUT_EXPECTED = [
    "Work Center","Order Type","Requirement Segment","Site","Sales Order","BTP Ticket",
    "Customer Contract ID","Sold-To PO No.","Status","Cost Category","Claim Cost",
    "Ship-To Party PO No.","Article","Cust Article No.","Article Lead Time","Model Name",
    "CRD","PD","PODD","LPD","Ship-To No.","Ship-To Search Term","Ship-To Name",
    "Ship-To Country","Document Date","Remark","Order Quantity",
    # contoh size (boleh lebih banyak, pola: UK_*)
    "UK_7-","UK_8","UK_8-","UK_9","UK_9-","UK_10","UK_10-","UK_11","UK_11-","UK_12-"
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
# Util khusus PIVOT UK_* → long
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
    }
    return alias_map.get(base, base)

def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [_normalize_header(c) for c in df.columns]
    return df

def _detect_size_cols(df: pd.DataFrame, prefix: str = "UK_") -> list[str]:
    return [c for c in df.columns if str(c).strip().upper().startswith(prefix)]

def pivot_one_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """Wide (kolom UK_*) -> Long (Size, Qty). Buang Qty NaN/0."""
    df = _normalize_columns(df)
    size_cols = _detect_size_cols(df, "UK_")
    if not size_cols:
        return pd.DataFrame()  # kembalikan kosong (akan diisi header FINAL_ORDER belakangan)

    id_vars = [c for c in df.columns if c not in size_cols]
    out = df.melt(id_vars=id_vars, value_vars=size_cols, var_name="Size", value_name="Qty")
    out["Qty"] = pd.to_numeric(out["Qty"], errors="coerce")
    out = out[out["Qty"].notna() & (out["Qty"] != 0)]
    return out

def finalize_columns(df_long: pd.DataFrame, ticket_date: str, factory_subject: str) -> pd.DataFrame:
    """Tambah kolom input user & kolom turunan, lalu urutkan sesuai FINAL_ORDER."""
    df = df_long.copy()
    df["Ticket Date"] = ticket_date
    df["Factory E-mail Subject"] = factory_subject
    df["Reduce Qty"] = df["Qty"]
    df["Increase Qty"] = 0
    df["New Qty"] = 0
    # Pastikan semua kolom target ada
    for col in FINAL_ORDER:
        if col not in df.columns:
            df[col] = pd.NA
    return df[FINAL_ORDER]

# --- Streamlit Page ---
set_page("PGD Apps — Input Tracking Report", "📝")
header("📝 Input Tracking Report")

# ==============
# Petunjuk Umum
# ==============
with st.expander("❓ Cara pakai (singkat) — klik untuk lihat"):
    st.markdown("""
**Tujuan:** menyiapkan file Excel untuk *tracking reroute*, *tracking PO Pending Cancel (filter+reorder)*, dan **pivot size UK_*** menjadi format Pending Cancel (Size, Qty).

**Format file:**
- Terima **.xlsx** / **.xls** (bisa multi-sheet).
- Nama kolom sebaiknya mengikuti tabel panduan di tiap sub-tools. Kolom yang hilang akan dibuat **NA** agar urutan tetap rapi.
- Data kolom lain tidak dihapus, kecuali pada mode **Pivot** hasil akhirnya disederhanakan ke kolom target.

**Langkah cepat:**
1) Pilih sub-tools.  
2) Upload Excel.  
3) (Khusus Pivot) isi **Ticket Date** & **Factory E-mail Subject**.  
4) Cek **Report/preview**, lalu **Download** hasil.
""")

choice = st.selectbox("Pilih sub-tools", [
    "Buat masukin trackingan PO Reroute",
    "Buat masukin tracking PO Pending Cancel",
    "Pivot Size → Pending Cancel Format (Size, Qty)"   # <— fitur baru
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
# 3) (Baru) Pivot Size → Pending Cancel Format
# =========================
elif choice == "Pivot Size → Pending Cancel Format (Size, Qty)":
    st.subheader("🧮 Pivot kolom UK_* → baris (Size, Qty) + Reorder ke Format Pending Cancel")
    st.caption("Tool ini mengubah kolom size `UK_*` menjadi baris `Size, Qty`, menambahkan kolom **Ticket Date** & **Factory E-mail Subject**, dan mengurutkan output sesuai format target.")

    with st.expander("📥 Kolom input yang direkomendasikan (contoh)"):
        st.dataframe(_df_from_list(PIVOT_INPUT_EXPECTED, "Kolom Input (contoh)"), use_container_width=True)

    with st.expander("📤 Kolom output & urutan akhir"):
        st.dataframe(_df_from_list(FINAL_ORDER, "Urutan Kolom Output"), use_container_width=True)

    # Input parameter global (untuk 1x proses semua sheet)
    c1, c2, c3 = st.columns([1,1,2])
    with c1:
        ticket_date = st.date_input("Ticket Date", help="Tanggal tiket untuk semua baris output")
    with c2:
        factory_subject = st.text_input("Factory E-mail Subject", value="Pending Cancel – Summary")
    with c3:
        st.markdown("")

    st.markdown("---")
    up3 = st.file_uploader("Upload Excel (.xlsx/.xls)", type=["xlsx","xls"], accept_multiple_files=False, key="itr_pivot")
    if not up3:
        st.stop()

    xls = pd.read_excel(up3, sheet_name=None, engine="openpyxl")
    out_frames = {}
    report_rows = []

    for sheet_name, df in xls.items():
        # 1) Pivot
        df_long = pivot_one_sheet(df)

        # 2) Jika tidak ada kolom UK_*, buat kosong dengan header final agar konsisten
        if df_long.empty or "Size" not in df_long.columns or "Qty" not in df_long.columns:
            empty_df = pd.DataFrame(columns=FINAL_ORDER)
            out_frames[sheet_name] = empty_df
            report_rows.append({
                "Sheet": sheet_name, "Status": "No UK_* columns", "Rows": 0,
                "Distinct Size": 0, "Total Qty": 0
            })
            continue

        # 3) Finalize (tambah kolom input + kolom turunan + reorder)
        final_df = finalize_columns(
            df_long,
            ticket_date.strftime("%Y-%m-%d"),
            factory_subject
        )
        out_frames[sheet_name] = final_df

        # 4) Report
        report_rows.append({
            "Sheet": sheet_name,
            "Status": "OK",
            "Rows": len(final_df),
            "Distinct Size": final_df["Size"].nunique(dropna=True),
            "Total Qty": pd.to_numeric(final_df["Qty"], errors="coerce").fillna(0).sum()
        })

    report_df = pd.DataFrame(report_rows).sort_values("Sheet")
    st.success("✅ Pivot selesai. Berikut ringkasannya:")
    st.dataframe(report_df, use_container_width=True)

    # Download
    payload = write_excel_autofit({**out_frames, "Report": report_df})
    st.download_button("⬇️ Download Hasil Pivot (Pending Cancel Format)", data=payload,
                       file_name="pivoted_output.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

footer()
