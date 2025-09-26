# ============================
# Input Tracking Report — with User Guide
# ============================

import pandas as pd
import os
import streamlit as st
from utils import set_page, header, footer, write_excel_autofit

# --- Urutan kolom yang diinginkan ---
DESIRED_ORDER = [
    "Working Status","Select","Working Status Descr.","Requirement Segment",
    "Sales Order","Sold-To PO No.","CRD","CRD-DRC","PD","POSDD-DRC","POSDD",
    "FPD","FPD-DRC","PODD","PODD-DRC","Est. Inspection Date",
    "LPD","LPD-DRC","FGR","Cust Article No.","Model Name","Article",
    "Lead Time","Season","Product Hierarchy 3","Ship-To Search Term",
    "Ship-To Country","Document Date","Order Quantity","Order Type"
]

# --- Kolom target & urutan yang diinginkan ---
TARGET_ORDER = [
    "Sales Order","Remark","Sold-To PO No.","Ship-To Party PO No.",
    "Model Name","Cust Article No.","Article","Order Quantity",
    "CRD","PD","LPD","PODD","Ship-To Search Term","Ship-To Name"
]

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

# --- Streamlit Page ---
set_page("PGD Apps — Input Tracking Report", "📝")
header("📝 Input Tracking Report")

# ==============
# Petunjuk Umum
# ==============
with st.expander("❓ Cara pakai (singkat) — klik untuk lihat"):
    st.markdown("""
**Tujuan:** membantu menyiapkan file Excel untuk kebutuhan *tracking reroute* dan *tracking PO Pending Cancel* agar kolomnya rapi & berurutan.

**Format file:**
- Terima **.xlsx** / **.xls** dengan **bisa multi-sheet**.
- Nama kolom harus sesuai tabel di bawah. Jika ada kolom yang tidak ada, tool akan otomatis membuat kolom kosong (NA) supaya urutannya tetap benar.
- Data lain di luar daftar kolom **tidak dihapus**, hanya dipindah ke belakang (untuk Reroute).

**Langkah cepat:**
1) Pilih sub-tools di bawah sesuai kebutuhan.  
2) Upload file Excel sumber.  
3) Cek **Report** untuk melihat kolom yang hilang/ditambahkan.  
4) Klik **Download** untuk ambil hasilnya.

**Tips kualitas data:**
- Pastikan **Sales Order / PO / tanggal** tidak berformat text yang aneh (spasi di depan/belakang).  
- Usahakan header kolom sesuai ejaan pada tabel (huruf besar-kecil tidak masalah, tool akan *normalize*).
""")

choice = st.selectbox("Pilih sub-tools", [
    "Buat masukin trackingan PO Reroute",
    "Buat masukin tracking PO Pending Cancel",
])

# =========================
# Guide: Kolom Wajib/Urutan
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

footer()
