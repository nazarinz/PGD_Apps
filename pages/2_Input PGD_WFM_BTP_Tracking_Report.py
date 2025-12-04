import re
from datetime import date, datetime
from typing import List, Optional
import io

import numpy as np
import pandas as pd
import streamlit as st

# =============== Page Configuration ===============
st.set_page_config(
    page_title="PO Tracking Normalizer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =============== Custom CSS for Better UI ===============
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .info-box {
        background-color: #e7f3ff;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
        margin: 1rem 0;
    }
    .success-box {
        background-color: #d4edda;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #28a745;
        margin: 1rem 0;
    }
    .warning-box {
        background-color: #fff3cd;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #ffc107;
        margin: 1rem 0;
    }
    .error-box {
        background-color: #f8d7da;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #dc3545;
        margin: 1rem 0;
    }
    .stButton>button {
        width: 100%;
        background-color: #1f77b4;
        color: white;
        font-weight: 600;
        padding: 0.5rem 1rem;
        border-radius: 0.5rem;
    }
    .stButton>button:hover {
        background-color: #155a8a;
    }
    </style>
""", unsafe_allow_html=True)


# =============== Helper Functions ===============
def normalize_input_columns_common(df: pd.DataFrame) -> pd.DataFrame:
    """Menyamakan nama kolom penting agar robust terhadap variasi header user."""
    def canon(s: str) -> str:
        s = (s or "").strip()
        s = re.sub(r"\uFEFF", "", s)
        s_low = s.lower()
        s_low = re.sub(r"[._:/\\\-]+", " ", s_low)
        s_low = re.sub(r"#", " #", s_low)
        s_low = re.sub(r"\s+", " ", s_low).strip()
        return s_low

    mapping = {
        "prod fact": "Work Center", "prod fact #": "Work Center", "work center": "Work Center",
        "so no": "Sales Order", "so": "Sales Order", "sales order": "Sales Order",
        "customer contract no": "Customer Contract ID", "customer contract id": "Customer Contract ID",
        "customer contract": "Customer Contract ID",
        "po": "Sold-To PO No.", "po no": "Sold-To PO No.", "po #": "Sold-To PO No.",
        "po number": "Sold-To PO No.", "sold to po no": "Sold-To PO No.",
        "sold to po number": "Sold-To PO No.",
        "ship to party po no": "Ship-To Party PO No.", "ship to party po number": "Ship-To Party PO No.",
        "change type": "Status", "status": "Status", "prod status": "Prod. Status",
        "cost type": "Cost Category", "cost category": "Cost Category",
        "crd": "CRD", "pd": "PD", "lpd": "LPD", "podd": "PODD",
        "art name": "Model Name", "model name": "Model Name",
        "art #": "Cust Article No.", "art": "Cust Article No.", "cust article no": "Cust Article No.",
        "cust article": "Cust Article No.",
        "article": "Article", "article lead time": "Article Lead Time",
        "cust #": "Ship-To Search Term", "cust": "Ship-To Search Term",
        "ship to search term": "Ship-To Search Term",
        "country": "Ship-To Country", "ship to country": "Ship-To Country",
        "ship-to country": "Ship-To Country", "ship to  country": "Ship-To Country",
        "document date": "Document Date", "doc date": "Document Date",
        "size": "Size", "ticket #": "Ticket", "ticket": "Ticket", "btp ticket": "BTP Ticket",
        "claim cost": "Claim Cost",
        "remark": "Remark", "remarks": "Remark",
        "qty": "Order Quantity", "order quantity": "Order Quantity", "order qty": "Order Quantity",
        "old quantity": "Old Quantity", "new quantity": "New Quantity",
        "reduce quantity": "Reduce", "reduce qty": "Reduce",
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


def _clean_money(x: str) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    return re.sub(r"[,$]", "", str(x).strip())


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


def _format_ticket_date_any(val) -> str:
    if val is None or val == "":
        return ""
    try:
        d = pd.to_datetime(val, errors="coerce")
    except Exception:
        return ""
    return d.strftime("%m/%d/%Y") if pd.notna(d) else ""


def _prefer_series(primary, secondary):
    """
    Pilih nilai dari primary kalau tidak kosong, kalau kosong pakai secondary.
    Dipakai untuk Cost Type: prefer 'Prod. Status', fallback 'Cost Category'.
    """
    if not isinstance(primary, pd.Series) or not isinstance(secondary, pd.Series):
        return primary
    a = primary.fillna("").astype(str).str.strip()
    b = secondary.fillna("").astype(str).str.strip()
    out = a.mask(a == "", b)
    out = out.replace({"nan": "", "None": ""})
    return out


# ✅ Tambahkan CRD sebelum LPD
TRACKING_COL_ORDER = [
    "Ticket Date", "Prod Fact.", "Document Date", "SO NO", "Customer Contract No",
    "PO#", "BTP Ticket", "Factory E-mail Subject", "Art.Name", "Art #", "Article",
    "Cust#", "Country", "Size", "Qty", "Reduce Qty", "Increase Qty", "New Qty",
    "CRD", "LPD", "PODD", "Change Type", "Cost Type", "Claim Cost",
]


# =============== 4A) NORMALIZER: CANCELLATION (AGGREGATED PER HEADER) ===============
def normalize_cancel_to_tracking(
    df_in: pd.DataFrame,
    ticket_date_val,
    subject_str: str,
    size_prefix: str = "UK_",
) -> tuple[pd.DataFrame, list]:
    """
    Mode: CANCELLATION (AGGREGATED)
    - TIDAK lagi per-size.
    - Kolom Size dikosongkan.
    - Qty & Reduce Qty = TOTAL Order Quantity per header (grand total).
    - Total diambil dari:
        1) Kolom 'Order Quantity' (jika ada & valid), kalau kosong/0 →
        2) Jumlah seluruh kolom size (UK_*) sebagai fallback.
    - Change Type = 'PO Cancelation'
    - Cost Type = dari 'Prod. Status' atau fallback 'Cost Category'
    """
    logs: list[str] = []
    logs.append("🔧 [CANCEL] Normalisasi header...")

    # 1) Normalisasi nama kolom umum
    df = normalize_input_columns_common(df_in)

    # 2) Drop kolom duplikat & all-null
    df = df.copy()
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.dropna(axis=1, how="all")

    # 3) Strip string & isi NaN jadi ""
    for c in df.columns:
        if df[c].dtype == "O":
            df[c] = df[c].astype(str).str.strip()
    df = df.fillna("")

    logs.append(f"✅ Kolom terdeteksi: {len(df.columns)} kolom")

    # 4) Deteksi kolom Remark
    if "Remark" in df.columns:
        remark_vals = df["Remark"].dropna().unique().tolist()
        logs.append(
            "✅ Kolom Remark ditemukan dengan nilai: "
            + ", ".join([str(v) for v in remark_vals[:10]])
        )
    else:
        logs.append("⚠️ Kolom Remark TIDAK ditemukan! Semua baris akan diperlakukan sama.")

    # 5) Deteksi kolom size (hanya untuk fallback total)
    size_cols = [c for c in df.columns if str(c).startswith(size_prefix)]
    logs.append(
        f"🔹 Kolom size terdeteksi: {len(size_cols)} kolom "
        f"({', '.join(size_cols[:5])}{'...' if len(size_cols) > 5 else ''})"
    )

    # 6) Filter baris untuk cancellation:
    if "Remark" in df.columns:
        logs.append("🔎 Filter baris Remark == 'Order Quantity' atau 'Cancel'")
        remark_clean = df["Remark"].astype(str).str.strip()
        use = df[remark_clean.isin(["Order Quantity", "Cancel"])].copy()
        logs.append(f"✅ Baris setelah filter Remark: {len(use)} baris")

        if use.empty:
            uniq_remark = (
                df["Remark"].astype(str).str.strip().replace("", "<blank>").unique().tolist()
            )
            logs.append("⚠️ Tidak ada baris dengan Remark = 'Order Quantity' atau 'Cancel'")
            logs.append(
                "⚠️ Nilai Remark yang ditemukan: "
                + ", ".join([str(v) for v in uniq_remark[:10]])
            )
            return pd.DataFrame(columns=TRACKING_COL_ORDER), logs
    else:
        logs.append("⚠️ Kolom 'Remark' tidak ada. Semua baris dipakai sebagai kandidat cancellation.")
        use = df.copy()

    if use.empty:
        logs.append("⚠️ Tidak ada baris yang bisa diproses setelah filter awal.")
        return pd.DataFrame(columns=TRACKING_COL_ORDER), logs

    # 7) Kolom header (key) per-header
    key_src = [
        "Work Center",
        "Document Date",
        "Sales Order",
        "Customer Contract ID",
        "Sold-To PO No.",
        "BTP Ticket",
        "Model Name",
        "Cust Article No.",
        "Article",
        "Ship-To Search Term",
        "Ship-To Country",
        "CRD",   # ✅ ikut sebagai key & output
        "LPD",
        "PODD",
        "Status",
        "Prod. Status",
        "Cost Category",
        "Claim Cost",
    ]
    key_cols = [c for c in key_src if c in use.columns]
    if not key_cols:
        logs.append("❌ Tidak ada kolom header kunci yang ditemukan. Periksa struktur file.")
        return pd.DataFrame(columns=TRACKING_COL_ORDER), logs

    # 8) Forward-fill info header
    use[key_cols] = use[key_cols].replace("", pd.NA).ffill().fillna("")

    # 9) Hitung total quantity per baris
    if "Order Quantity" in use.columns:
        use["_OrderQty_from_col"] = pd.to_numeric(use["Order Quantity"], errors="coerce")
    else:
        use["_OrderQty_from_col"] = np.nan
        logs.append("ℹ️ Kolom 'Order Quantity' tidak tersedia, akan coba hitung dari size (UK_*) saja.")

    if size_cols:
        use["_OrderQty_from_sizes"] = (
            use[size_cols]
            .apply(lambda r: pd.to_numeric(r, errors="coerce").fillna(0).sum(), axis=1)
        )
    else:
        use["_OrderQty_from_sizes"] = np.nan
        logs.append("ℹ️ Kolom size tidak tersedia, fallback dari size tidak bisa digunakan.")

    # 10) Agregasi per header
    logs.append(f"🔗 Grouping per header dengan key: {', '.join(key_cols)}")
    grp = (
        use.groupby(key_cols, dropna=False)[["_OrderQty_from_col", "_OrderQty_from_sizes"]]
        .agg({
            "_OrderQty_from_col": lambda s: s.dropna().iloc[0] if s.dropna().size else np.nan,
            "_OrderQty_from_sizes": "max",
        })
        .reset_index()
    )

    # 11) Tentukan total quantity final per header
    grp["TotalQty"] = grp["_OrderQty_from_col"]
    mask_fallback = grp["TotalQty"].isna() | (grp["TotalQty"] == 0)
    grp.loc[mask_fallback, "TotalQty"] = grp.loc[mask_fallback, "_OrderQty_from_sizes"]

    logs.append(
        f"✅ TotalQty berhasil dihitung untuk {grp['TotalQty'].notna().sum()} header "
        f"(fallback dari size dipakai untuk {mask_fallback.sum()} header)."
    )

    # 12) Format tanggal & Claim Cost
    if "Document Date" in grp.columns:
        grp["Document Date"] = _fmt_shortdate_series(grp["Document Date"])
    if "CRD" in grp.columns:
        grp["CRD"] = _fmt_shortdate_series(grp["CRD"])
    if "LPD" in grp.columns:
        grp["LPD"] = _fmt_shortdate_series(grp["LPD"])
    if "PODD" in grp.columns:
        grp["PODD"] = _fmt_shortdate_series(grp["PODD"])
    if "Claim Cost" in grp.columns:
        grp["Claim Cost"] = grp["Claim Cost"].apply(
            lambda x: (f"${_to_float(x):,.2f}" if pd.notna(_to_float(x)) else "")
        )

    # 13) Bangun output final sesuai template tracking
    out = pd.DataFrame(index=grp.index)

    ticket_date_str = _format_ticket_date_any(ticket_date_val)
    subject_str = (subject_str or "").strip()

    out["Ticket Date"] = ticket_date_str
    out["Prod Fact."] = grp.get("Work Center", "")
    out["Document Date"] = grp.get("Document Date", "")
    out["SO NO"] = grp.get("Sales Order", "")
    out["Customer Contract No"] = grp.get("Customer Contract ID", "")
    out["PO#"] = grp.get("Sold-To PO No.", "")
    out["BTP Ticket"] = grp.get("BTP Ticket", "")
    out["Factory E-mail Subject"] = subject_str
    out["Art.Name"] = grp.get("Model Name", "")
    out["Art #"] = grp.get("Cust Article No.", "")
    out["Article"] = grp.get("Article", "")
    out["Cust#"] = grp.get("Ship-To Search Term", "")
    out["Country"] = grp.get("Ship-To Country", "")

    # Size kosong untuk cancellation
    out["Size"] = ""

    out["Qty"] = grp["TotalQty"]
    out["Reduce Qty"] = grp["TotalQty"]
    out["Increase Qty"] = ""
    out["New Qty"] = ""

    # ✅ CRD sebelum LPD
    out["CRD"] = grp.get("CRD", "")
    out["LPD"] = grp.get("LPD", "")
    out["PODD"] = grp.get("PODD", "")

    # Change Type fixed by mode
    out["Change Type"] = "PO Cancelation"

    # Cost Type = Prod. Status → Cost Category (fallback)
    prod_series = grp.get("Prod. Status", pd.Series([""] * len(grp)))
    costcat_series = grp.get("Cost Category", pd.Series([""] * len(grp)))
    out["Cost Type"] = _prefer_series(prod_series, costcat_series)

    out["Claim Cost"] = grp.get("Claim Cost", "")

    # Format angka
    for col in ["Qty", "Reduce Qty", "Increase Qty", "New Qty"]:
        if col in out.columns:
            as_float = pd.to_numeric(out[col], errors="coerce")
            out[col] = as_float.apply(
                lambda v: "" if pd.isna(v)
                else (str(int(v)) if float(v).is_integer() else f"{v}")
            )

    out = out.reindex(columns=TRACKING_COL_ORDER)

    logs.append(f"✅ Berhasil memproses {len(out)} baris data (1 baris per header, no per-size).")
    return out, logs


# =============== 4B) NORMALIZER: QUANTITY CHANGE (PER SIZE + BTP TICKET) ===============
def normalize_quantity_to_tracking(
    df_in: pd.DataFrame,
    ticket_date_val,
    subject_str: str,
    size_prefix: str = "UK_",
) -> tuple[pd.DataFrame, list]:
    """
    Mode: QUANTITY CHANGE (PER SIZE)
    - Gunakan Remark = Ticket / Old Quantity / New Quantity / Reduce
    - Ticket diambil dari baris Remark = "Ticket" (UK_* berisi nomor ticket per size)
    - Change Type = 'PO Quantity change'
    - Cost Type = Prod. Status atau fallback Cost Category
    """
    logs: list[str] = []
    logs.append("🔧 [QTY CHANGE] Normalisasi header...")

    df = normalize_input_columns_common(df_in)

    df = df.copy()
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.dropna(axis=1, how="all")

    for c in df.columns:
        if df[c].dtype == "O":
            df[c] = df[c].astype(str).str.strip()
    df = df.fillna("")

    logs.append(f"✅ Kolom terdeteksi: {len(df.columns)} kolom")

    # Debug Remark
    if "Remark" in df.columns:
        remark_vals = df["Remark"].dropna().unique().tolist()
        logs.append(f"✅ Kolom Remark ditemukan dengan nilai: {', '.join([str(v) for v in remark_vals[:10]])}")
    else:
        logs.append(f"⚠️ Kolom Remark TIDAK ditemukan! Kolom yang ada: {', '.join([str(c) for c in df.columns[:10]])}")

    size_cols = [c for c in df.columns if str(c).startswith(size_prefix)]
    logs.append(f"🔹 Kolom size terdeteksi: {len(size_cols)} kolom ({', '.join(size_cols[:5])}{'...' if len(size_cols) > 5 else ''})")

    if not size_cols:
        raise ValueError(f"Tidak ditemukan kolom size yang diawali '{size_prefix}'")

    if "Remark" not in df.columns:
        raise ValueError("Kolom 'Remark' tidak ditemukan setelah normalisasi. Periksa nama kolom di file Excel Anda.")

    # --------- 1) Build ticket_map dari Remark = "Ticket" ---------
    ticket_map: dict[str, str] = {}
    ffill_cols_base = [
        "Work Center", "Document Date", "Sales Order", "Customer Contract ID",
        "Sold-To PO No.", "Model Name", "Cust Article No.", "Article",
        "Ship-To Search Term", "Ship-To Country", "CRD", "LPD", "PODD",
        "Status", "Prod. Status", "Cost Category", "Claim Cost"
    ]
    ffill_cols_base = [c for c in ffill_cols_base if c in df.columns]

    df_ticket = df[df["Remark"] == "Ticket"].copy()
    if not df_ticket.empty:
        logs.append(f"🔎 Ditemukan baris Remark = 'Ticket': {len(df_ticket)} baris")
        if ffill_cols_base:
            df_ticket[ffill_cols_base] = df_ticket[ffill_cols_base].replace("", pd.NA).ffill().fillna("")
        long_ticket = df_ticket.melt(
            id_vars=ffill_cols_base,
            value_vars=size_cols,
            var_name="Size",
            value_name="Ticket_raw"
        )
        long_ticket["Ticket_raw"] = long_ticket["Ticket_raw"].astype(str).str.strip()
        long_ticket = long_ticket[long_ticket["Ticket_raw"] != ""]
        if not long_ticket.empty:
            long_ticket["_key"] = long_ticket[ffill_cols_base + ["Size"]].astype(str).agg("|".join, axis=1)
            ticket_map = dict(zip(long_ticket["_key"], long_ticket["Ticket_raw"]))
            logs.append(f"✅ Ticket map terbentuk untuk {len(ticket_map)} kombinasi header+size")
        else:
            logs.append("⚠️ Baris Ticket ada, tapi semua nilai size kosong. Ticket map kosong.")
    else:
        logs.append("ℹ️ Tidak ada baris Remark = 'Ticket'. BTP Ticket akan dikosongkan.")

    # --------- 2) Proses Old/New/Reduce ---------
    use = df[df["Remark"].isin(["Old Quantity", "New Quantity", "Reduce"])].copy()
    logs.append(f"✅ Baris dengan Remark Old/New/Reduce: {len(use)} baris")

    if use.empty:
        logs.append("⚠️ Tidak ada baris Old Quantity / New Quantity / Reduce")
        return pd.DataFrame(columns=TRACKING_COL_ORDER), logs

    ffill_cols = ffill_cols_base
    if ffill_cols:
        use[ffill_cols] = use[ffill_cols].replace("", pd.NA).ffill().fillna("")

    long = use.melt(
        id_vars=ffill_cols + ["Remark"],
        value_vars=size_cols,
        var_name="Size",
        value_name="Qty_raw"
    )

    long["Qty"] = pd.to_numeric(long["Qty_raw"], errors="coerce")
    long = long[long["Qty"].notna()]

    if long.empty:
        logs.append("⚠️ Tidak ada data quantity setelah unpivot")
        return pd.DataFrame(columns=TRACKING_COL_ORDER), logs

    pivot = long.pivot_table(
        index=ffill_cols + ["Size"],
        columns="Remark",
        values="Qty",
        aggfunc="first"
    ).reset_index()

    # --------- 3) Mapping BTP Ticket ke pivot per header+size ---------
    if ticket_map:
        pivot["_key"] = pivot[ffill_cols + ["Size"]].astype(str).agg("|".join, axis=1)
        pivot["BTP_Ticket_Mapped"] = pivot["_key"].map(ticket_map).fillna("")
        logs.append("✅ BTP Ticket berhasil dipetakan ke pivot per size")
    else:
        pivot["BTP_Ticket_Mapped"] = ""

    # Format tanggal & Claim Cost
    if "Document Date" in pivot.columns:
        pivot["Document Date"] = _fmt_shortdate_series(pivot["Document Date"])
    if "CRD" in pivot.columns:
        pivot["CRD"] = _fmt_shortdate_series(pivot["CRD"])
    if "LPD" in pivot.columns:
        pivot["LPD"] = _fmt_shortdate_series(pivot["LPD"])
    if "PODD" in pivot.columns:
        pivot["PODD"] = _fmt_shortdate_series(pivot["PODD"])
    if "Claim Cost" in pivot.columns:
        pivot["Claim Cost"] = pivot["Claim Cost"].apply(
            lambda x: (f"${_to_float(x):,.2f}" if pd.notna(_to_float(x)) else "")
        )

    # Bangun output
    out = pd.DataFrame(index=pivot.index)
    ticket_date_str = _format_ticket_date_any(ticket_date_val)
    subject_str = (subject_str or "").strip()

    out["Ticket Date"] = ticket_date_str
    out["Prod Fact."] = pivot.get("Work Center", "")
    out["Document Date"] = pivot.get("Document Date", "")
    out["SO NO"] = pivot.get("Sales Order", "")
    out["Customer Contract No"] = pivot.get("Customer Contract ID", "")
    out["PO#"] = pivot.get("Sold-To PO No.", "")
    out["BTP Ticket"] = pivot.get("BTP_Ticket_Mapped", "")
    out["Factory E-mail Subject"] = subject_str
    out["Art.Name"] = pivot.get("Model Name", "")
    out["Art #"] = pivot.get("Cust Article No.", "")
    out["Article"] = pivot.get("Article", "")
    out["Cust#"] = pivot.get("Ship-To Search Term", "")
    out["Country"] = pivot.get("Ship-To Country", "")
    out["Size"] = pivot.get("Size", "")

    old_q = pivot.get("Old Quantity")
    new_q = pivot.get("New Quantity")
    red_q = pivot.get("Reduce")

    out["Qty"] = old_q

    old_f = pd.to_numeric(old_q, errors="coerce")
    new_f = pd.to_numeric(new_q, errors="coerce")
    red_f = pd.to_numeric(red_q, errors="coerce")

    inc = np.full(len(pivot), np.nan)
    red = red_f.copy()

    if red_f.isna().all():
        red = np.where((~pd.isna(old_f)) & (~pd.isna(new_f)) & (new_f < old_f),
                       old_f - new_f, np.nan)
        inc = np.where((~pd.isna(old_f)) & (~pd.isna(new_f)) & (new_f > old_f),
                       new_f - old_f, np.nan)

    out["Reduce Qty"] = red
    out["Increase Qty"] = inc
    out["New Qty"] = new_q

    # ✅ CRD sebelum LPD
    out["CRD"] = pivot.get("CRD", "")
    out["LPD"] = pivot.get("LPD", "")
    out["PODD"] = pivot.get("PODD", "")

    # Change Type fixed by mode
    out["Change Type"] = "PO Quantity change"

    # Cost Type = Prod. Status → Cost Category (fallback)
    prod_series = pivot.get("Prod. Status", pd.Series([""] * len(pivot)))
    costcat_series = pivot.get("Cost Category", pd.Series([""] * len(pivot)))
    out["Cost Type"] = _prefer_series(prod_series, costcat_series)

    out["Claim Cost"] = pivot.get("Claim Cost", "")

    for col in ["Qty", "Reduce Qty", "Increase Qty", "New Qty"]:
        if col in out.columns:
            as_float = pd.to_numeric(out[col], errors="coerce")
            out[col] = as_float.apply(
                lambda v: "" if pd.isna(v) else (str(int(v)) if float(v).is_integer() else f"{v}")
            )

    out = out.reindex(columns=TRACKING_COL_ORDER)
    logs.append(
        f"✅ Berhasil memproses {len(out)} baris data "
        f"(per size, BTP Ticket terisi kalau ada baris Remark='Ticket')"
    )
    return out, logs


# =============== Main Streamlit App ===============
def main():
    # Header
    st.markdown('<p class="main-header">📊 PO Tracking Normalizer</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Convert Quantity Change & Cancellation Reports to Tracking Format</p>', unsafe_allow_html=True)

    # Sidebar - Instructions
    with st.sidebar:
        st.image("https://via.placeholder.com/300x100/1f77b4/ffffff?text=PO+Tracker", use_container_width=True)
        st.markdown("### 📖 Panduan Penggunaan")
        st.markdown("""
        **Langkah-langkah:**
        1. Upload file Excel (.xlsx)
        2. Pilih jenis tiket
        3. Isi informasi tambahan
        4. Klik "Process File"
        5. Download hasil
        
        **Format File:**
        - Harus berformat Excel (.xlsx)
        - Harus memiliki kolom size (UK_X) untuk Quantity Change
        - Dianjurkan memiliki kolom Remark
        
        **Nilai Remark yang Valid:**
        - **Cancellation:** `Cancel` atau `Order Quantity`
        - **Quantity Change:** `Ticket`, `Old Quantity`, `New Quantity`, `Reduce`
        
        **Change Type (output):**
        - Quantity Change → `PO Quantity change`
        - Cancellation → `PO Cancelation`
        
        **Cost Type (output):**
        - Mengambil dari `Prod. Status`, jika kosong → `Cost Category`
        """)
        st.markdown("---")
        st.markdown("### ℹ️ Info")
        st.info("Tool ini akan mengkonversi file PO menjadi format tracking standar dengan normalisasi kolom otomatis.")

    # Main content area
    col1, col2 = st.columns([2, 1])

    with col1:
        st.markdown("### 📁 Upload File")
        uploaded_file = st.file_uploader(
            "Pilih file Excel",
            type=['xlsx'],
            help="Upload file Excel yang berisi data PO Quantity Change atau Cancellation"
        )

    with col2:
        st.markdown("### ⚙️ Pengaturan")
        mode = st.radio(
            "Jenis Tiket:",
            options=["Quantity Change", "Cancellation"],
            help="Pilih jenis tiket sesuai dengan data yang akan diproses"
        )

    # Additional inputs
    st.markdown("### 📝 Informasi Tambahan")
    col3, col4 = st.columns(2)

    with col3:
        ticket_date = st.date_input(
            "Ticket Date",
            value=datetime.now(),
            help="Tanggal tiket dibuat"
        )

    with col4:
        email_subject = st.text_input(
            "Factory E-mail Subject",
            placeholder="Masukkan subject email...",
            help="Subject dari email factory terkait tiket ini"
        )

    # Advanced settings (collapsible)
    with st.expander("🔧 Pengaturan Lanjutan"):
        size_prefix = st.text_input(
            "Size Column Prefix",
            value="UK_",
            help="Prefix untuk kolom size (default: UK_)"
        )

    st.markdown("---")

    if uploaded_file is not None:
        st.markdown('<div class="info-box">', unsafe_allow_html=True)
        st.markdown(f"**File:** {uploaded_file.name}")
        st.markdown(f"**Size:** {uploaded_file.size / 1024:.2f} KB")
        st.markdown('</div>', unsafe_allow_html=True)

        if st.button("🚀 Process File", type="primary", use_container_width=True):
            try:
                with st.spinner("⏳ Memproses file..."):
                    df_in = pd.read_excel(uploaded_file, dtype=str)

                    # Preview
                    with st.expander("👀 Preview Data Input"):
                        st.dataframe(df_in.head(10), use_container_width=True)
                        st.caption(f"Menampilkan 10 dari {len(df_in)} baris")

                        if 'Remark' in df_in.columns:
                            st.markdown("**Nilai kolom Remark yang terdeteksi:**")
                            remark_values = df_in['Remark'].dropna().unique().tolist()
                            st.code(", ".join([str(v) for v in remark_values[:10]]))

                            remark_set = set([str(v).strip() for v in remark_values])
                            cancel_indicators = {"Cancel", "Order Quantity"}
                            qty_change_indicators = {"Ticket", "Old Quantity", "New Quantity", "Reduce"}

                            detected_mode = None
                            if remark_set & cancel_indicators:
                                detected_mode = "Cancellation"
                            elif remark_set & qty_change_indicators:
                                detected_mode = "Quantity Change"

                            if detected_mode and detected_mode != mode:
                                st.warning(
                                    f"⚠️ Mode terdeteksi: **{detected_mode}** "
                                    f"(Anda memilih: **{mode}**). Pertimbangkan untuk mengganti mode."
                                )
                            elif detected_mode:
                                st.success(f"✅ Mode yang dipilih sesuai dengan data: **{mode}**")

                    # Process
                    if mode == "Cancellation":
                        st.info(f"🔄 Memproses sebagai: **{mode} (aggregated per header)**")
                        result, logs = normalize_cancel_to_tracking(
                            df_in,
                            ticket_date_val=ticket_date,
                            subject_str=email_subject,
                            size_prefix=size_prefix
                        )
                        output_prefix = "tracking_cancel"
                    else:
                        st.info(f"🔄 Memproses sebagai: **{mode} (per size)**")
                        result, logs = normalize_quantity_to_tracking(
                            df_in,
                            ticket_date_val=ticket_date,
                            subject_str=email_subject,
                            size_prefix=size_prefix
                        )
                        output_prefix = "tracking_qtychange"

                    # Logs
                    with st.expander("📋 Processing Logs"):
                        for log in logs:
                            st.text(log)

                    if result.empty:
                        st.markdown('<div class="warning-box">', unsafe_allow_html=True)
                        st.warning("⚠️ Tidak ada data yang dihasilkan.")
                        st.markdown('</div>', unsafe_allow_html=True)

                        st.markdown("### 💡 Kemungkinan Penyebab:")
                        if mode == "Cancellation":
                            st.markdown("""
                            - Pastikan kolom **Remark** berisi nilai `Cancel` atau `Order Quantity`
                            - Jika tidak ada kolom Remark, pastikan struktur header sudah benar
                            - Jika ingin pakai fallback size, pastikan ada kolom `UK_*` berisi qty
                            """)
                        else:
                            st.markdown("""
                            - Pastikan ada baris **Remark = 'Ticket'** untuk BTP Ticket per size (opsional)
                            - Pastikan kolom **Remark** berisi nilai `Old Quantity`, `New Quantity`, atau `Reduce`
                            - Pastikan ada kolom size yang diawali dengan `UK_` dan ada nilai qty
                            """)

                        st.info("📋 Lihat **Processing Logs** di atas untuk detail lebih lanjut")
                    else:
                        st.markdown('<div class="success-box">', unsafe_allow_html=True)
                        st.success(f"✅ Berhasil! {len(result)} baris data telah diproses")
                        st.markdown('</div>', unsafe_allow_html=True)

                        st.markdown("### 📊 Preview Hasil")
                        st.dataframe(result, use_container_width=True)

                        st.markdown("### 💾 Download Hasil")

                        col5, col6, col7 = st.columns([1, 1, 1])

                        with col5:
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                result.to_excel(writer, index=False)
                            output.seek(0)
                            output_name = f"{output_prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                            st.download_button(
                                label="📥 Download Excel",
                                data=output,
                                file_name=output_name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )

                        with col6:
                            csv = result.to_csv(index=False)
                            csv_name = f"{output_prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
                            st.download_button(
                                label="📥 Download CSV",
                                data=csv,
                                file_name=csv_name,
                                mime="text/csv",
                                use_container_width=True
                            )

                        with col7:
                            st.metric("Total Rows", len(result))

                        with st.expander("📈 Statistik Data"):
                            col8, col9, col10 = st.columns(3)

                            with col8:
                                unique_so = result['SO NO'].nunique()
                                st.metric("Unique SO", unique_so)

                            with col9:
                                unique_articles = result['Art #'].nunique()
                                st.metric("Unique Articles", unique_articles)

                            with col10:
                                total_qty = pd.to_numeric(result['Qty'], errors='coerce').sum()
                                st.metric("Total Qty", f"{int(total_qty):,}")

            except Exception as e:
                st.markdown('<div class="error-box">', unsafe_allow_html=True)
                st.error(f"❌ Error: {str(e)}")
                st.markdown('</div>', unsafe_allow_html=True)

                with st.expander("🔍 Detail Error"):
                    st.exception(e)

    else:
        st.markdown('<div class="info-box">', unsafe_allow_html=True)
        st.info("👆 Silakan upload file Excel untuk memulai proses")
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")
    st.markdown(
        "<p style='text-align: center; color: #666;'>Made with ❤️ for PO Tracking | © 2024</p>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
