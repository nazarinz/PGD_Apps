# pages/2_Quantity Change Normalizer.py
# ② Normalizer (Excel) — Reshape UK_* + Ticket Date & Subject
# Adapted from user's PO Tools (Normalizer Only)

import io
import re
from datetime import datetime, date
from typing import List, Optional

import numpy as np
import pandas as pd
import streamlit as st
from utils import set_page, header, footer

# =========================================================
#                   PAGE HEADER
# =========================================================
set_page("PGD Apps — Quantity Change Normalizer", "🧾")
header("🧾 Quantity Change Tools — ② Normalizer (Excel)")

# =========================================================
#              NORMALIZER — HELPERS & MAPPINGS
# =========================================================
def normalize_input_columns(df: pd.DataFrame) -> pd.DataFrame:
    def canon(s: str) -> str:
        s = (s or "").strip()
        s = re.sub(r"\uFEFF", "", s)
        s_low = s.lower()
        s_low = re.sub(r"[._:/\\\-]+", " ", s_low)
        s_low = re.sub(r"#", " #", s_low)
        s_low = re.sub(r"\s+", " ", s_low).strip()
        return s_low

    mapping = {
        "prod fact": "Work Center",
        "prod fact #": "Work Center",
        "work center": "Work Center",

        "so no": "Sales Order",
        "so": "Sales Order",
        "sales order": "Sales Order",

        "customer contract no": "Customer Contract ID",
        "customer contract id": "Customer Contract ID",
        "customer contract": "Customer Contract ID",

        "po": "Sold-To PO No.",
        "po no": "Sold-To PO No.",
        "po #": "Sold-To PO No.",
        "po number": "Sold-To PO No.",
        "sold to po no": "Sold-To PO No.",
        "sold to po number": "Sold-To PO No.",

        "ship to party po no": "Ship-To Party PO No.",
        "ship to party po number": "Ship-To Party PO No.",

        "change type": "Status",
        "status": "Status",

        "cost type": "Cost Category",
        "cost category": "Cost Category",

        "crd": "CRD",
        "pd": "PD",
        "lpd": "LPD",
        "podd": "PODD",

        "art name": "Model Name",
        "model name": "Model Name",

        "art #": "Cust Article No.",
        "art": "Cust Article No.",
        "cust article no": "Cust Article No.",
        "cust article": "Cust Article No.",

        "article": "Article",
        "article lead time": "Article Lead Time",

        "cust #": "Ship-To Search Term",
        "cust": "Ship-To Search Term",
        "ship to search term": "Ship-To Search Term",

        "country": "Ship-To Country",
        "ship to country": "Ship-To Country",
        "ship-to country": "Ship-To Country",
        "ship to  country": "Ship-To Country",

        "document date": "Document Date",
        "doc date": "Document Date",

        "size": "Size",

        "ticket #": "Ticket",
        "ticket": "Ticket",

        "claim cost": "Claim Cost",

        "qty": "Old Quantity",
        "old qty": "Old Quantity",
        "new qty": "New Quantity",
        "reduce qty": "Reduce",
        "reduce": "Reduce",

        # tambahan alias umum
        "order quantity": "Order Quantity",
        "order qty": "Order Quantity",
        "old quantity": "Old Quantity",
        "new quantity": "New Quantity",
        "reduce quantity": "Reduce",
    }

    rename_map = {}
    for col in df.columns:
        if str(col).startswith("UK_"):
            continue
        key = canon(str(col))
        target = mapping.get(key)
        if target:
            rename_map[col] = target

    df2 = df.copy().rename(columns=rename_map)
    df2.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df2.columns]
    return df2


def rename_output_columns(df: pd.DataFrame) -> pd.DataFrame:
    out_map = {
        "Work Center": "Prod Fact.",
        "Sales Order": "SO NO",
        "Customer Contract ID": "Customer Contract No",
        "Sold-To PO No.": "PO#",
        "Status": "Change Type",
        "Cost Category": "Cost Type",
        "Model Name": "Art.Name",
        "Cust Article No.": "Art #",
        "Ship-To Search Term": "Cust#",
        "Ship-To Country": "Country",
        "Ticket": "Ticket#",
        "Old Quantity": "Qty",
        "New Quantity": "New Qty",
        "Reduce": "Reduce Qty",
    }
    return df.rename(columns=out_map)


def _clean_money(x: str) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    x = str(x).strip()
    return re.sub(r"[,$]", "", x)


def _to_float(x):
    x = _clean_money(x)
    if x == "":
        return np.nan
    try:
        return float(x)
    except Exception:
        return np.nan


def _fmt_shortdate_series(s: pd.Series) -> pd.Series:
    dt = pd.to_datetime(s, errors="coerce")
    out = dt.dt.strftime("%m/%d/%Y")
    return out.mask(dt.isna(), "")


# ====================== FIXED COLS (2 MODE) ======================
# Quantity change
FIXED_COLS_QTYCHANGE = [
    "Select",
    "Status",
    "Working Status",
    "Working Status Descr.",
    "PO Date",
    "Requirement Segment",
    "Order Type",
    "Site",
    "Work Center",
    "Sales Order",
    "Sold-To PO No.",
    "Cost Category",
    "Feedback Status",
    "Feedback Date",
    "Ship-To Party PO No.",
    "CRD",
    "PD",
    "Prod. Team ATP",
    "FPD",
    "FPD-DRC",
    "POSDD",
    "POSDD-DRC",
    "LPD",
    "LPD-DRC",
    "PODD",
    "PODD-DRC",
    "FGR",
    "Model Name",
    "Cust Article No.",
    "Part Number",
    "Gender",
    "Article",
    "Article Lead Time",
    "Develop Type",
    "Last Code",
    "Season",
    "Product Hierarchy 3",
    "Outsole Mold",
    "Pattern Code (Upper",
    "Ship-To No.",
    "Ship-To Search Term",
    "Ship-To Name",
    "Ship-To Country",
    "Shipping Type",
    "Packing Type",
    "VAS Cut-Off Date",
    "Classification Code",
    "Changed By",
    "Document Date",
    "Remark",
    "Order Quantity",
]

# Cancellation
FIXED_COLS_CANCEL = [
    "Select",
    "Status",
    "Working Status",
    "Working Status Descr.",
    "PO Date",
    "Requirement Segment",
    "Order Type",
    "Site",
    "Work Center",
    "Sales Order",
    "Customer Contract ID",
    "BTP Ticket",
    "Sold-To PO No.",
    "Prod. Status",
    "Claim Cost",
    "Ship-To Party PO No.",
    "CRD",
    "PD",
    "Prod. Team ATP",
    "FPD",
    "FPD-DRC",
    "POSDD",
    "POSDD-DRC",
    "LPD",
    "LPD-DRC",
    "PODD",
    "PODD-DRC",
    "FGR",
    "Model Name",
    "Cust Article No.",
    "Part Number",
    "Gender",
    "Article",
    "Article Lead Time",
    "Develop Type",
    "Last Code",
    "Season",
    "Product Hierarchy 3",
    "Outsole Mold",
    "Pattern Code (Upper",
    "Ship-To No.",
    "Ship-To Search Term",
    "Ship-To Name",
    "Ship-To Country",
    "Shipping Type",
    "Packing Type",
    "VAS Cut-Off Date",
    "Classification Code",
    "Changed By",
    "Document Date",
    "Remark",
    "Order Quantity",
]


# =========================================================
#                  CORE RESHAPE LOGIC
# =========================================================
def reshape_po(
    df: pd.DataFrame,
    fixed_cols_all: List[str] = None,
    size_prefix: str = "UK_",
) -> pd.DataFrame:
    if fixed_cols_all is None:
        fixed_cols_all = [
            "Work Center", "Sales Order", "Customer Contract ID", "Sold-To PO No.",
            "Ship-To Party PO No.", "Status", "Cost Category", "CRD", "PD", "LPD", "PODD",
            "Model Name", "Cust Article No.", "Article", "Article Lead Time",
            "Ship-To Search Term", "Ship-To Country", "Document Date", "Remark", "Order Quantity",
        ]

    # PATCH: remove duplicate columns (e.g. Shipping Type dua kali)
    df = df.copy()
    df = df.loc[:, ~df.columns.duplicated()]

    df = df.dropna(axis=1, how="all")
    for c in df.columns:
        if df[c].dtype == "O":
            df[c] = df[c].astype(str).str.strip()
    df = df.fillna("")

    fixed_cols = [c for c in fixed_cols_all if c in df.columns]
    size_cols = [c for c in df.columns if str(c).startswith(size_prefix)]
    if not size_cols:
        raise ValueError(
            f"Tidak ditemukan kolom size yang diawali '{size_prefix}' "
            f"(mis. UK_5K, UK_1-, dst)."
        )

    # forward fill untuk kolom fixed (kecuali Remark, Order Quantity)
    ffill_cols = [c for c in fixed_cols if c not in ("Remark", "Order Quantity")]
    if ffill_cols:
        df[ffill_cols] = df[ffill_cols].replace("", pd.NA).ffill().fillna("")

    group_key_cols = [c for c in fixed_cols if c not in ("Remark", "Order Quantity")]
    if not group_key_cols:
        raise ValueError("Kolom kunci tidak ditemukan. Pastikan header sesuai.")
    df["_group_key"] = df[group_key_cols].apply(
        lambda r: "|".join([str(v) for v in r.values]), axis=1
    )

    # Ticket & Claim Cost per group+size (Remark = 'Ticket' / 'Claim Cost')
    ticket_map, claim_map = {}, {}
    for _, row in df.iterrows():
        remark = row.get("Remark", "")
        if remark not in ("Ticket", "Claim Cost"):
            continue
        gkey = row["_group_key"]
        for sc in size_cols:
            val = str(row.get(sc, "")).strip()
            if val == "" or val.lower() in ("nan", "none"):
                continue
            if remark == "Ticket":
                ticket_map[(gkey, sc)] = val
            else:
                claim_map[(gkey, sc)] = val

    if "Remark" not in df.columns:
        raise ValueError("Kolom 'Remark' tidak ada di data.")
    use = df[df["Remark"].isin({"Old Quantity", "New Quantity", "Reduce"})].copy()

    long = use.melt(
        id_vars=group_key_cols + ["Remark"],
        value_vars=size_cols,
        var_name="Size",
        value_name="Qty",
    )
    long = long[long["Qty"].astype(str).str.strip() != ""]

    pivot = long.pivot_table(
        index=group_key_cols + ["Size"],
        columns="Remark",
        values="Qty",
        aggfunc="first",
    ).reset_index()

    def _map_ticket(row):
        gk = "|".join([str(row[c]) for c in group_key_cols])
        return ticket_map.get((gk, row["Size"]), "")

    def _map_claim(row):
        gk = "|".join([str(row[c]) for c in group_key_cols])
        return claim_map.get((gk, row["Size"]), "")

    pivot["Ticket"] = pivot.apply(_map_ticket, axis=1)
    pivot["Claim Cost"] = pivot.apply(_map_claim, axis=1)

    # Buang baris yang tidak punya Ticket
    pivot = pivot[pivot["Ticket"].astype(str).str.strip() != ""].copy()

    # Konversi angka
    for c in ["Old Quantity", "New Quantity", "Reduce"]:
        if c in pivot.columns:
            pivot[c] = pivot[c].apply(_to_float)

    # Format Claim Cost sebagai string uang
    if "Claim Cost" in pivot.columns:
        pivot["Claim Cost"] = pivot["Claim Cost"].apply(
            lambda x: (f"${_to_float(x):,.2f}" if pd.notna(_to_float(x)) else "")
        )

    # Format beberapa tanggal
    for dc in ["CRD", "PD", "LPD", "PODD", "Document Date"]:
        if dc in pivot.columns:
            pivot[dc] = _fmt_shortdate_series(pivot[dc])

    # Bersihkan Customer Contract ID kosong / nol
    if "Customer Contract ID" in pivot.columns:
        col = pivot["Customer Contract ID"].astype(str).str.strip()
        mask_empty = (col.isin({"", "nan", "NaN", "None"}) | col.str.fullmatch(r"0+"))
        pivot["Customer Contract ID"] = col.mask(mask_empty, "")

    pivot["Ticket"] = (
        pivot["Ticket"].astype(str).str.strip().replace({"nan": "", "None": ""})
    )

    std_order = [
        "Work Center", "Sales Order", "Customer Contract ID", "Sold-To PO No.",
        "Ship-To Party PO No.", "Status", "Cost Category", "CRD", "PD", "LPD", "PODD",
        "Model Name", "Cust Article No.", "Article", "Article Lead Time",
        "Ship-To Search Term", "Ship-To Country", "Document Date",
        "Size", "Ticket", "Claim Cost", "Old Quantity", "New Quantity", "Reduce",
    ]
    std_order = [c for c in std_order if c in pivot.columns] + \
                [c for c in pivot.columns if c not in std_order]
    out = pivot[std_order].copy()

    sort_cols = [
        c for c in ["Work Center", "Sales Order", "Model Name", "Cust Article No.", "Size"]
        if c in out.columns
    ]
    if sort_cols:
        out = out.sort_values(sort_cols, kind="mergesort")

    return out.reset_index(drop=True)


# ===================== FINAL ORDER (COMMON) =====================
FINAL_ORDER = [
    "Ticket Date",
    "Prod Fact.",
    "Document Date",
    "SO NO",
    "Customer Contract No",
    "PO#",
    "BTP Ticket",
    "Factory E-mail Subject",
    "Art.Name",
    "Art #",
    "Article",
    "Cust#",
    "Country",
    "Size",
    "Qty",
    "Reduce Qty",
    "Increase Qty",
    "New Qty",
    "LPD",
    "PODD",
    "Change Type",
    "Cost Type",
    "Claim Cost",
    "Final Status",
    "Leftover Qty (FG)",
    "Final cost",
    "Remark",
    "Cancel/ Update Date",
    "Result",
    "Propose Check FG (Y/N)",
    "Remark2",
    "Email",
]


def _format_ticket_date_any(val) -> str:
    """Terima string/date/datetime/Timestamp -> 'MM/DD/YYYY' atau '' jika invalid."""
    if val is None or val == "":
        return ""
    try:
        d = pd.to_datetime(val, errors="coerce")
    except Exception:
        return ""
    return d.strftime("%m/%d/%Y") if pd.notna(d) else ""


def add_fixed_fields_and_select(
    df_out: pd.DataFrame,
    ticket_date_val,
    subject_str: str,
) -> pd.DataFrame:
    df_out = df_out.copy()

    # 1) Ticket Date
    df_out["Ticket Date"] = _format_ticket_date_any(ticket_date_val)

    # 2) Subject
    df_out["Factory E-mail Subject"] = (subject_str or "").strip()

    # 3) Increase Qty (jika New Qty > Qty)
    qty = pd.to_numeric(df_out.get("Qty"), errors="coerce")
    newq = pd.to_numeric(df_out.get("New Qty"), errors="coerce")
    inc = np.where(
        (~pd.isna(qty)) & (~pd.isna(newq)) & (newq > qty),
        newq - qty,
        np.nan,
    )
    df_out["Increase Qty"] = inc

    # 4) Kolom tambahan yang mungkin belum ada → buat kosong
    extra_cols = [
        "BTP Ticket",
        "Final Status",
        "Leftover Qty (FG)",
        "Final cost",
        "Remark",
        "Cancel/ Update Date",
        "Result",
        "Propose Check FG (Y/N)",
        "Remark2",
        "Email",
    ]
    for c in extra_cols:
        if c not in df_out.columns:
            df_out[c] = ""

    # 5) Susun kolom final (hanya yang tersedia)
    have_cols = [c for c in FINAL_ORDER if c in df_out.columns]
    df_out = df_out[have_cols].copy()

    # 6) Format angka
    for col in ["Qty", "Reduce Qty", "Increase Qty", "New Qty", "Leftover Qty (FG)"]:
        if col in df_out.columns:
            as_float = pd.to_numeric(df_out[col], errors="coerce")
            df_out[col] = as_float.apply(
                lambda v: "" if pd.isna(v)
                else (str(int(v)) if float(v).is_integer() else f"{v}")
            )

    return df_out


# =========================================================
#                   STREAMLIT UI
# =========================================================
st.subheader("② Normalizer (Excel) — Reshape UK_* + Ticket Date & Subject")

st.markdown(
    """
**Cara pakai singkat:**
1) Pilih jenis tiket (**Quantity Change** / **Cancellation**).  
2) Upload Excel sumber (sheet pertama).  
3) Isi **Ticket Date** (sekali untuk semua baris) dan **Factory E-mail Subject**.  
4) Klik **Proses & Download**.  

**Catatan:** Aplikasi otomatis menangkap **semua kolom yang diawali `UK_`** (robust untuk variasi ukuran).
"""
)

mode = st.radio("Jenis Ticket", ["Quantity Change", "Cancellation"], index=0)

file_xlsx = st.file_uploader(
    "Upload Excel (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=False,
)

colA, colB, colC = st.columns([1, 1, 1])
with colA:
    tdate: Optional[date] = st.date_input(
        "Ticket Date", value=None, format="MM/DD/YYYY"
    )
with colB:
    subj = st.text_input("Factory E-mail Subject", value="")
with colC:
    clicked = st.button("⚙️ Proses & Download")

if clicked:
    if file_xlsx is None:
        st.error("Silakan upload file Excel dulu.")
    elif tdate is None:
        st.error("Silakan isi Ticket Date.")
    else:
        try:
            df_in = pd.read_excel(file_xlsx, sheet_name=0, dtype=str)
            df_in = normalize_input_columns(df_in)

            if mode == "Quantity Change":
                fixed_cols = FIXED_COLS_QTYCHANGE
            else:
                fixed_cols = FIXED_COLS_CANCEL

            hasil_std = reshape_po(df_in, fixed_cols_all=fixed_cols)
            hasil_lbl = rename_output_columns(hasil_std)

            # langsung pass objek tdate (date)
            hasil_final = add_fixed_fields_and_select(hasil_lbl, tdate, subj)

            st.success(f"Sukses! {len(hasil_final):,} baris dihasilkan.")
            st.dataframe(hasil_final, use_container_width=True)

            bio = io.BytesIO()
            with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
                hasil_final.to_excel(writer, index=False, sheet_name="Result")
                ws = writer.sheets["Result"]
                for i, c in enumerate(hasil_final.columns):
                    # PATCH: aman untuk DataFrame kosong
                    col_values = hasil_final[c].astype(str).tolist()
                    if col_values:
                        max_content = max(len(v) for v in col_values)
                        max_len = max(len(str(c)), max_content)
                    else:
                        max_len = len(str(c))
                    ws.set_column(i, i, min(max(10, max_len + 2), 50))
                ws.freeze_panes(1, 0)

            st.download_button(
                "⬇️ Download Excel (Normalizer)",
                data=bio.getvalue(),
                file_name=f"hasil_konversi_PO_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.exception(e)

footer()
