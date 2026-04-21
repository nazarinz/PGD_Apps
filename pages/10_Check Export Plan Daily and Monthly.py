from utils.auth import require_login

require_login()

# pages/2_Check Export Plan Daily and Monthly.py
# PGD Apps — Check Export Plan Daily and Monthly (SO Auto-Detect) + Drill-down Matches (cached)

import io
import re
import logging
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Set

import pandas as pd
import streamlit as st

# =============================================================================
# Logger
# =============================================================================
logging.basicConfig(
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# =============================================================================
# Page config
# =============================================================================
st.set_page_config(
    page_title="PGD Apps — Check Export Plan Daily and Monthly",
    page_icon="🔎",
    layout="wide",
)
st.title("🔎 SO Auto-Detect — Export Plan Daily and Monthly")

# =============================================================================
# Constants
# =============================================================================

# Kolom yang dianggap sebagai "anchor" untuk mendeteksi baris header
HEADER_ANCHOR_COLS: Set[str] = {
    "no", "invoice", "fvb so", "fvb_so", "fvbso",
    "sto_dn", "sto dn", "sap_cont_no", "sap cont no",
    "sap_odr_no", "sap odr no", "sales order", "salesorder",
    "cust", "customer", "dest", "dest code", "art. no", "art no",
    "qty", "ctn", "cbm", "line", "remarks",
}

# Kandidat nama kolom SO (SAP)
SO_CANDIDATES: List[str] = [
    "sap_odr_no", "sap_odrno", "sap_order_no", "sap_odr_number",
    "sales_order", "salesorder", "so", "so_no", "sono", "so_number",
]

# Kandidat nama kolom FVB SO
FVB_CANDIDATES: List[str] = [
    "fvb_so", "fvbso", "fvb so", "fvb_sales_order", "fvb",
]

MAX_HEADER_SCAN_ROWS: int = 20   # Batas baris scan untuk deteksi header
HEADER_ANCHOR_MIN_MATCH: int = 2  # Minimal kolom anchor yang harus cocok

# =============================================================================
# File reading utilities
# =============================================================================

def _read_excel_bytes(file_bytes: bytes, suffix: str) -> Dict[str, pd.DataFrame]:
    """Baca bytes Excel/ODS → dict {sheet_name: DataFrame}."""
    b = io.BytesIO(file_bytes)
    sfx = suffix.lower()
    if sfx in (".xlsx", ".xlsm"):
        return pd.read_excel(b, engine="openpyxl", sheet_name=None, header=None, dtype=object)
    elif sfx == ".xls":
        try:
            return pd.read_excel(io.BytesIO(file_bytes), engine="xlrd", sheet_name=None, header=None, dtype=object)
        except Exception:
            try:
                return pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl", sheet_name=None, header=None, dtype=object)
            except Exception as e:
                raise RuntimeError("Gagal membaca .xls. Pastikan file tidak korup atau misnamed.") from e
    elif sfx == ".xlsb":
        try:
            return pd.read_excel(b, engine="pyxlsb", sheet_name=None, header=None, dtype=object)
        except Exception as e:
            raise RuntimeError("Gagal membaca .xlsb. Install 'pyxlsb' atau konversi ke .xlsx.") from e
    elif sfx == ".ods":
        try:
            return pd.read_excel(b, engine="odf", sheet_name=None, header=None, dtype=object)
        except Exception as e:
            raise RuntimeError("Gagal membaca ODS. Install 'odfpy' atau konversi ke .xlsx.") from e
    else:
        raise ValueError(f"Format tidak didukung: {sfx}")


def read_any_raw(uploaded_file) -> Dict[str, pd.DataFrame]:
    """
    Baca uploaded_file → dict {sheet_name: DataFrame} tanpa header parsing.
    Semua format didukung: CSV, XLSX, XLS, XLSB, XLSM, ODS.
    """
    name = uploaded_file.name
    suf = name.strip().lower()

    if suf.endswith(".csv"):
        try:
            uploaded_file.seek(0)
        except Exception:
            pass
        df = pd.read_csv(uploaded_file, header=None, dtype=object)
        return {"Sheet1": df}
    elif suf.endswith((".xlsx", ".xls", ".ods", ".xlsb", ".xlsm")):
        if hasattr(uploaded_file, "getvalue"):
            raw_bytes = uploaded_file.getvalue()
        else:
            try:
                uploaded_file.seek(0)
            except Exception:
                pass
            raw_bytes = uploaded_file.read()
        ext = "." + name.split(".")[-1].lower()
        return _read_excel_bytes(raw_bytes, ext)
    else:
        raise ValueError("Format tidak didukung. Gunakan CSV/XLSX/XLS/XLSB/XLSM/ODS.")


# =============================================================================
# Header row auto-detection
# =============================================================================

def detect_header_row_index(df_raw: pd.DataFrame) -> Tuple[int, int]:
    """
    Scan baris awal untuk menemukan baris header berdasarkan HEADER_ANCHOR_COLS.
    Kembalikan (header_row_index, match_count).
    Default ke baris 0 jika tidak ada yang cocok.
    """
    best_idx = 0
    best_score = 0

    for i in range(min(MAX_HEADER_SCAN_ROWS, len(df_raw))):
        row_vals = [
            str(x).strip().lower()
            for x in df_raw.iloc[i].tolist()
            if pd.notna(x) and str(x).strip()
        ]
        score = sum(1 for v in row_vals if v in HEADER_ANCHOR_COLS)
        if score > best_score:
            best_score = score
            best_idx = i
        if score >= HEADER_ANCHOR_MIN_MATCH:
            # Kalau sudah cukup yakin, langsung pakai
            if score >= 3:
                break

    logger.info(f"Header row detected at index {best_idx} (score={best_score})")
    return best_idx, best_score


def apply_header_row(df_raw: pd.DataFrame, header_row_index: int) -> pd.DataFrame:
    """
    Terapkan baris header_row_index sebagai nama kolom,
    lalu ambil data dari baris berikutnya ke bawah.
    """
    if header_row_index >= len(df_raw):
        return df_raw

    raw_header = [
        str(x).strip() if (x is not None and str(x).strip() not in ("None", "nan", "")) else f"col_{ci}"
        for ci, x in enumerate(df_raw.iloc[header_row_index].tolist())
    ]

    # Deduplikasi nama kolom
    seen: Dict[str, int] = {}
    uniq_cols = []
    for c in raw_header:
        if c in seen:
            seen[c] += 1
            uniq_cols.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            uniq_cols.append(c)

    df2 = df_raw.iloc[header_row_index + 1:].copy()
    df2.columns = uniq_cols
    df2 = df2.dropna(how="all").dropna(axis=1, how="all")
    df2 = df2.reset_index(drop=True)
    return df2


def load_sheet_with_auto_header(df_raw: pd.DataFrame) -> Tuple[pd.DataFrame, int, int]:
    """
    Auto-detect header row dan kembalikan (df_clean, header_row_index, match_score).
    """
    hdr_idx, score = detect_header_row_index(df_raw)
    df_clean = apply_header_row(df_raw, hdr_idx)
    return df_clean, hdr_idx, score


# =============================================================================
# Column normalization helpers
# =============================================================================

def norm_col_name(c: str) -> str:
    c = str(c).strip().lower()
    c = re.sub(r"\s+", " ", c)
    c = re.sub(r"[^0-9a-z_ ]+", "", c)
    c = c.replace(" ", "_")
    c = re.sub(r"_+", "_", c)
    return c.strip("_")


def normalize_columns(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """
    Normalisasi nama kolom → {norm_name: original_name}.
    Kembalikan (df_dengan_kolom_ternorm, mapping).
    """
    mapping: Dict[str, str] = {}
    new_cols: List[str] = []
    seen: Set[str] = set()
    for c in df.columns:
        nc = norm_col_name(str(c)) or "col"
        base = nc
        counter = 1
        while base in seen:
            counter += 1
            base = f"{nc}_{counter}"
        seen.add(base)
        mapping[base] = str(c)
        new_cols.append(base)
    out = df.copy()
    out.columns = new_cols
    return out, mapping


def pick_column(df: pd.DataFrame, target_name: str) -> Optional[str]:
    """Cari kolom di df dengan nama persis / case-insensitive / normalized."""
    if df is None or df.empty:
        return None
    for c in df.columns:
        if str(c) == target_name:
            return str(c)
    tl = target_name.lower()
    for c in df.columns:
        if str(c).lower() == tl:
            return str(c)
    nt = norm_col_name(target_name)
    for c in df.columns:
        if norm_col_name(c) == nt:
            return str(c)
    return None


# =============================================================================
# SO column detection
# =============================================================================

def detect_so_column(df: pd.DataFrame) -> Tuple[Optional[str], str]:
    """
    Deteksi kolom SAP SO (SAP_ODR_NO / Sales Order / SO) dari DataFrame.
    Kembalikan (original_col_name, reason).
    """
    if df is None or df.empty or df.dropna(how="all").empty:
        return None, "DataFrame kosong"

    df_n, mapping = normalize_columns(df)
    cols = list(df_n.columns)

    # 1) Exact candidate
    for cand in SO_CANDIDATES:
        nc = norm_col_name(cand)
        if nc in cols:
            return mapping[nc], f"Exact candidate match: {nc}"

    # 2) Heuristic: sap + odr/order
    for cn in cols:
        if "sap" in cn and ("odr" in cn or "order" in cn):
            return mapping[cn], f"Heuristic: sap+odr/order → {cn}"

    # 3) Heuristic: sales + order
    for cn in cols:
        if "sales" in cn and "order" in cn:
            return mapping[cn], f"Heuristic: sales+order → {cn}"

    # 4) Fallback short names
    for cn in cols:
        if cn in {"order", "so"}:
            return mapping[cn], f"Fallback exact short name: {cn}"

    return None, "Tidak ditemukan kolom SO yang cocok"


def detect_fvb_so_column(df: pd.DataFrame) -> Tuple[Optional[str], str]:
    """Deteksi kolom FVB SO dari DataFrame. Kembalikan (original_col_name, reason)."""
    if df is None or df.empty:
        return None, "DataFrame kosong"

    df_n, mapping = normalize_columns(df)
    cols = list(df_n.columns)

    for cand in FVB_CANDIDATES:
        nc = norm_col_name(cand)
        if nc in cols:
            return mapping[nc], f"Exact FVB candidate match: {nc}"

    for cn in cols:
        if "fvb" in cn:
            return mapping[cn], f"Heuristic: contains 'fvb' → {cn}"

    return None, "Kolom FVB SO tidak ditemukan"


# =============================================================================
# SO value normalization
# =============================================================================

def normalize_so_series(s: pd.Series) -> pd.Series:
    """
    Normalisasi nilai SO agar bisa dibandingkan lintas file/format:
    - Handle float/int types sebelum konversi string
    - Strip semua whitespace (termasuk non-breaking space)
    - Hapus trailing .0 / .00 / .000 (float artifact dari Excel/ODS)
    - Ganti nan/none/'' dengan pd.NA
    """
    def _safe_str(v) -> str:
        if pd.isna(v):
            return ""
        if isinstance(v, float):
            if v == int(v):
                return str(int(v))
            return str(v)
        if isinstance(v, int):
            return str(v)
        return str(v)

    s2 = s.map(_safe_str)
    s2 = s2.str.strip().str.replace(r"\s+", "", regex=True)
    s2 = s2.str.replace(r"\.0+$", "", regex=True)
    s2 = s2.replace({"nan": pd.NA, "none": pd.NA, "None": pd.NA, "NaT": pd.NA, "": pd.NA})
    return s2


# =============================================================================
# Preferred sheet selection
# =============================================================================

PREFERRED_SHEET_NAMES: List[str] = ["loading_plan", "loading plan", "sheet1"]

def pick_best_sheet(sheets: Dict[str, pd.DataFrame]) -> str:
    """Pilih sheet terbaik dari dict. Prioritaskan nama yang ada di PREFERRED_SHEET_NAMES."""
    keys_lower = {k.strip().lower().replace(" ", "_"): k for k in sheets}
    for pref in PREFERRED_SHEET_NAMES:
        if pref in keys_lower:
            return keys_lower[pref]
    return next(iter(sheets.keys()))


# =============================================================================
# Result dataclass
# =============================================================================

@dataclass
class SOResult:
    base_file_name: str
    base_sheet_name: str
    base_header_row: int
    base_header_score: int
    # Kolom SO #1 (SAP_ODR_NO / fallback)
    base_col: str
    base_reason: str
    # Kolom SO #2 (FVB SO)
    base_col_fvb: Optional[str]
    base_reason_fvb: str
    ref_count: int
    matches: pd.DataFrame
    not_found: pd.DataFrame
    empty_so: pd.DataFrame
    base_df_out: pd.DataFrame       # berisi __SO_norm_SAP__ dan __SO_norm_FVB__
    ref_tables: List[Dict]
    log_df: pd.DataFrame
    matched_so_list: List[str]      # SO yang benar-benar match di referensi


# =============================================================================
# Core processing
# =============================================================================

def process_files(base_file, ref_files) -> SOResult:
    # ── 1. Baca base file tanpa header ──────────────────────────────────────
    base_sheets_raw = read_any_raw(base_file)
    base_sheet_name = pick_best_sheet(base_sheets_raw)
    base_df_raw = base_sheets_raw[base_sheet_name]

    # ── 2. Auto-detect header row ────────────────────────────────────────────
    base_df, hdr_idx, hdr_score = load_sheet_with_auto_header(base_df_raw)
    logger.info(
        f"Base sheet '{base_sheet_name}': header at row {hdr_idx}, "
        f"score={hdr_score}, shape={base_df.shape}"
    )

    # ── 3. Detect SO column #1 (SAP_ODR_NO) ─────────────────────────────────
    base_col = pick_column(base_df, "SAP_ODR_NO")
    if base_col:
        base_reason = f"Forced exact: SAP_ODR_NO | header=baris {hdr_idx + 1}"
    else:
        base_col, base_reason = detect_so_column(base_df)
        if base_col:
            base_reason = f"{base_reason} | header=baris {hdr_idx + 1}"

    # PENTING: Jangan fallback ke kolom pertama — kolom pertama sering berisi
    # nomor urut (NO), bukan SO number. Biarkan None jika tidak ditemukan.
    # Matching akan tetap jalan via FVB SO column.
    if base_col is None:
        base_reason = f"Tidak ditemukan kolom SAP SO | header=baris {hdr_idx + 1}"

    # ── 4. Detect SO column #2 (FVB SO) ──────────────────────────────────────
    base_col_fvb = pick_column(base_df, "FVB SO")
    if base_col_fvb:
        base_reason_fvb = f"Forced exact: FVB SO | header=baris {hdr_idx + 1}"
    else:
        base_col_fvb, base_reason_fvb = detect_fvb_so_column(base_df)
        if base_col_fvb:
            base_reason_fvb = f"{base_reason_fvb} | header=baris {hdr_idx + 1}"

    # ── 5. Tambah kolom normalisasi ───────────────────────────────────────────
    base_df_out = base_df.copy()
    if base_col and base_col in base_df_out.columns:
        base_df_out["__SO_norm_SAP__"] = normalize_so_series(base_df_out[base_col])
    else:
        base_df_out["__SO_norm_SAP__"] = pd.NA

    if base_col_fvb and base_col_fvb in base_df_out.columns:
        base_df_out["__SO_norm_FVB__"] = normalize_so_series(base_df_out[base_col_fvb])
    else:
        base_df_out["__SO_norm_FVB__"] = pd.NA

    # ── 6. Kumpulkan SO dari ref files ────────────────────────────────────────
    all_ref_sos: Set[str] = set()
    so_to_files: Dict[str, Set[str]] = {}
    log_rows: List[Tuple[str, str, str]] = []
    ref_tables: List[Dict] = []

    for f in (ref_files or []):
        try:
            sheets_raw = read_any_raw(f)
        except Exception as e:
            log_rows.append((f.name, "—", f"Gagal membaca file: {e}"))
            continue

        for sh, df_raw_ref in sheets_raw.items():
            # Auto-detect header di ref file juga
            try:
                df_ref, ref_hdr_idx, ref_hdr_score = load_sheet_with_auto_header(df_raw_ref)
            except Exception as e:
                log_rows.append((f.name, sh, f"Gagal parse header: {e}"))
                continue

            if df_ref.empty:
                log_rows.append((f.name, sh, "Sheet kosong setelah header detection"))
                continue

            col_ref, reason_ref = detect_so_column(df_ref)
            log_rows.append((
                f.name, sh,
                f"{reason_ref} | header=baris {ref_hdr_idx + 1} (score={ref_hdr_score})"
            ))

            if not col_ref or col_ref not in df_ref.columns:
                log_rows.append((f.name, sh, f"Kolom '{col_ref}' tidak valid setelah cleanup"))
                continue

            df_ref["__SO_norm__"] = normalize_so_series(df_ref[col_ref])
            sos_valid = df_ref["__SO_norm__"].dropna()

            if sos_valid.empty:
                log_rows.append((f.name, sh, f"Kolom '{col_ref}' tidak punya nilai SO valid"))
                continue

            for so in sos_valid.unique().tolist():
                all_ref_sos.add(so)
                so_to_files.setdefault(so, set()).add(f.name)

            ref_tables.append({
                "file": f.name,
                "sheet": sh,
                "so_col": col_ref,
                "header_row": ref_hdr_idx,
                "df": df_ref,
            })

    # ── 7. Matching: SAP atau FVB ─────────────────────────────────────────────
    out = base_df_out.copy()

    def _found(row) -> bool:
        sap = row["__SO_norm_SAP__"]
        fvb = row["__SO_norm_FVB__"]
        return (pd.notna(sap) and sap in all_ref_sos) or \
               (pd.notna(fvb) and fvb in all_ref_sos)

    def _source_files(row) -> str:
        files: Set[str] = set()
        sap = row["__SO_norm_SAP__"]
        fvb = row["__SO_norm_FVB__"]
        if pd.notna(sap) and sap in so_to_files:
            files |= so_to_files[sap]
        if pd.notna(fvb) and fvb in so_to_files:
            files |= so_to_files[fvb]
        return ", ".join(sorted(files)) if files else ""

    out["Found_in_reference"] = out.apply(_found, axis=1)
    out["Source_Files"] = out.apply(_source_files, axis=1)

    matches = out[out["Found_in_reference"]].copy()
    not_found = out[
        (~out["Found_in_reference"]) &
        (out["__SO_norm_SAP__"].notna() | out["__SO_norm_FVB__"].notna())
    ].copy()
    empty_so = out[
        out["__SO_norm_SAP__"].isna() & out["__SO_norm_FVB__"].isna()
    ].copy()

    # SO yang benar-benar ada di referensi (gabungan kedua kolom)
    matched_sap = set(matches["__SO_norm_SAP__"].dropna().astype(str))
    matched_fvb = set(matches["__SO_norm_FVB__"].dropna().astype(str))
    matched_so_list = sorted((matched_sap | matched_fvb) & all_ref_sos)

    log_df = (
        pd.DataFrame(log_rows, columns=["File", "Sheet", "Reason"])
        if log_rows
        else pd.DataFrame(columns=["File", "Sheet", "Reason"])
    )

    logger.info(
        f"Result — matches={len(matches)}, not_found={len(not_found)}, "
        f"empty_so={len(empty_so)}, matched_so_list={len(matched_so_list)}"
    )

    return SOResult(
        base_file_name=base_file.name,
        base_sheet_name=base_sheet_name,
        base_header_row=hdr_idx,
        base_header_score=hdr_score,
        base_col=base_col,
        base_reason=base_reason,
        base_col_fvb=base_col_fvb,
        base_reason_fvb=base_reason_fvb,
        ref_count=len(ref_files or []),
        matches=matches,
        not_found=not_found,
        empty_so=empty_so,
        base_df_out=base_df_out,
        ref_tables=ref_tables,
        log_df=log_df,
        matched_so_list=matched_so_list,
    )


# =============================================================================
# UI — Upload section
# =============================================================================

col1, col2 = st.columns(2)
with col1:
    base_file = st.file_uploader(
        "📂 Daily file (CSV/XLSX/XLS/ODS/XLSB/XLSM)",
        type=["csv", "xlsx", "xls", "ods", "xlsb", "xlsm"],
        accept_multiple_files=False,
    )
with col2:
    ref_files = st.file_uploader(
        "📂 Reference files (bisa banyak)",
        type=["csv", "xlsx", "xls", "ods", "xlsb", "xlsm"],
        accept_multiple_files=True,
    )

colA, colB = st.columns([1, 1])
with colA:
    do_process = st.button("▶️ Proses", type="primary")
with colB:
    do_reset = st.button("♻️ Reset hasil")

if do_reset:
    st.session_state.pop("so_detect_result", None)
    st.success("Hasil direset. Silakan pilih file dan klik Proses lagi.")

if do_process:
    if not base_file:
        st.error("⚠️ Unggah dulu **Daily file** sebelum memproses.")
        st.stop()
    with st.spinner("Memproses file…"):
        try:
            result = process_files(base_file, ref_files)
            st.session_state["so_detect_result"] = result
            st.success("✅ Proses selesai!")
        except Exception as e:
            st.error(f"❌ Gagal memproses: {e}")
            logger.exception("process_files failed")
            st.stop()

# =============================================================================
# UI — Render results (dari cache)
# =============================================================================

if "so_detect_result" in st.session_state:
    res: SOResult = st.session_state["so_detect_result"]

    # ── Info dasar file ───────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("📄 Daily file (base)")

    info_cols = st.columns(3)
    info_cols[0].info(f"**File:** {res.base_file_name}")
    info_cols[1].info(f"**Sheet:** {res.base_sheet_name}")
    info_cols[2].info(
        f"**Header terdeteksi:** baris ke-{res.base_header_row + 1} "
        f"(skor={res.base_header_score})"
    )

    col_info1, col_info2 = st.columns(2)
    with col_info1:
        if res.base_col:
            st.success(f"**Kolom SO #1 (SAP):** `{res.base_col}`\n\n_{res.base_reason}_")
        else:
            st.warning(f"**Kolom SO #1 (SAP):** Tidak ditemukan\n\n_{res.base_reason}_")
    with col_info2:
        if res.base_col_fvb:
            st.success(f"**Kolom SO #2 (FVB):** `{res.base_col_fvb}`\n\n_{res.base_reason_fvb}_")
        else:
            st.warning(f"**Kolom SO #2 (FVB):** Tidak ditemukan\n\n_{res.base_reason_fvb}_")

    # ── Statistik ringkas ─────────────────────────────────────────────────────
    st.markdown("---")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("✅ Matches", len(res.matches))
    m2.metric("❌ Not Found", len(res.not_found))
    m3.metric("⚠️ Empty SO", len(res.empty_so))
    m4.metric("🔗 Unique SO Match", len(res.matched_so_list))

    # ── Tabel utama ───────────────────────────────────────────────────────────
    st.subheader("✅ Matches")
    st.dataframe(res.matches.head(2000), use_container_width=True)

    st.subheader("❌ Not Found")
    st.dataframe(res.not_found.head(2000), use_container_width=True)

    st.subheader("⚠️ Empty SO")
    st.dataframe(res.empty_so.head(2000), use_container_width=True)

    if not res.log_df.empty:
        with st.expander("🧾 Log Deteksi Kolom (Ref Files)", expanded=False):
            st.dataframe(res.log_df, use_container_width=True)

    # ── DEBUG DIAGNOSTIC ──────────────────────────────────────────────────────
    with st.expander("🛠️ Debug: Cek nilai SO aktual (buka jika 0 matches)", expanded=False):
        st.markdown("**Nilai FVB SO di base file (sample 20):**")
        fvb_vals = res.base_df_out["__SO_norm_FVB__"].dropna().unique()[:20]
        st.write(list(fvb_vals))

        st.markdown("**Nilai SAP SO di base file (sample 20):**")
        sap_vals = res.base_df_out["__SO_norm_SAP__"].dropna().unique()[:20]
        st.write(list(sap_vals))

        st.markdown("**SO dari ref files (sample 30):**")
        for item in res.ref_tables:
            ref_vals = item["df"]["__SO_norm__"].dropna().unique()[:30]
            st.write(f"**{item['file']} · {item['sheet']} (kolom: {item['so_col']})**")
            st.write(list(ref_vals))

        # Cek apakah ada yang mirip tapi tidak exact match
        st.markdown("**Perbandingan karakter (FVB SO base vs ref — 3 sample pertama):**")
        all_ref_flat = []
        for item in res.ref_tables:
            all_ref_flat += item["df"]["__SO_norm__"].dropna().unique().tolist()
        all_ref_flat = list(set(all_ref_flat))

        for bv in list(fvb_vals)[:3]:
            st.write(f"Base FVB: `{repr(bv)}` (len={len(str(bv))})")
            # Cari yang paling mirip di ref
            candidates = [rv for rv in all_ref_flat if str(rv)[:6] == str(bv)[:6]]
            for rv in candidates[:3]:
                st.write(f"  Ref candidate: `{repr(rv)}` (len={len(str(rv))}) — equal={bv==rv}")

    # ── Drill-down ────────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("🔎 Detail SO yang Match (Drill-down)")

    if res.matched_so_list:
        show_all = st.checkbox("Tampilkan semua SO match", value=True)
        selected_sos: List[str] = (
            res.matched_so_list
            if show_all
            else st.multiselect(
                "Pilih SO yang ingin ditampilkan:",
                options=res.matched_so_list,
                default=res.matched_so_list[:1],
            )
        )

        if selected_sos:
            # Opsional: filter tabel Matches
            with st.expander("Filter tabel Matches ke SO terpilih (opsional)", expanded=False):
                if st.checkbox("Aktifkan filter Matches"):
                    mask = (
                        res.matches["__SO_norm_SAP__"].isin(selected_sos) |
                        res.matches["__SO_norm_FVB__"].isin(selected_sos)
                    )
                    st.dataframe(res.matches[mask], use_container_width=True)

            for so in selected_sos:
                st.markdown(f"### SO: **{so}**")

                # Baris di base yang cocok via SAP atau FVB
                mask_base = (
                    (res.base_df_out["__SO_norm_SAP__"] == so) |
                    (res.base_df_out["__SO_norm_FVB__"] == so)
                )
                base_rows = res.base_df_out[mask_base]

                matched_via: List[str] = []
                if (res.base_df_out["__SO_norm_SAP__"] == so).any():
                    matched_via.append(f"SAP_ODR_NO (`{res.base_col}`)")
                if (res.base_df_out["__SO_norm_FVB__"] == so).any():
                    matched_via.append(f"FVB SO (`{res.base_col_fvb}`)")
                if matched_via:
                    st.caption(f"✅ Ditemukan via: {' & '.join(matched_via)}")

                st.write("**Daily (base) — Rows**")
                if base_rows.empty:
                    st.info("Tidak ada baris di base untuk SO ini.")
                else:
                    st.dataframe(base_rows, use_container_width=True)

                # Baris di referensi
                any_ref_found = False
                for item in res.ref_tables:
                    sub = item["df"][item["df"]["__SO_norm__"] == so]
                    if not sub.empty:
                        any_ref_found = True
                        st.write(
                            f"**Reference — {item['file']} · {item['sheet']} "
                            f"(kolom SO: {item['so_col']} | header baris {item['header_row'] + 1})**"
                        )
                        st.dataframe(sub, use_container_width=True)
                if not any_ref_found:
                    st.warning("SO ini tidak ditemukan di tabel referensi yang tersimpan.")
        else:
            st.info("Tidak ada SO terpilih.")
    else:
        st.info("Belum ada SO yang match — detail tidak tersedia.")

    # ── Download Excel ────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("📥 Download Report (Excel)")

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"SO_AutoDetect_Report_{ts}.xlsx"
    output_buffer = io.BytesIO()

    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        res.matches.to_excel(writer, index=False, sheet_name="Matches")
        res.not_found.to_excel(writer, index=False, sheet_name="Not_Found")
        res.empty_so.to_excel(writer, index=False, sheet_name="Empty_SO")
        if not res.log_df.empty:
            res.log_df.to_excel(writer, index=False, sheet_name="Detection_Log")
        meta = pd.DataFrame({
            "Key": [
                "Base file", "Base sheet",
                "Header row (0-based)", "Header detection score",
                "SO column #1 (SAP)", "Reason #1",
                "SO column #2 (FVB)", "Reason #2",
                "Ref files count", "Generated at",
            ],
            "Value": [
                res.base_file_name, res.base_sheet_name,
                res.base_header_row, res.base_header_score,
                res.base_col or "(manual)", res.base_reason,
                res.base_col_fvb or "(tidak ditemukan)", res.base_reason_fvb,
                res.ref_count, ts,
            ],
        })
        meta.to_excel(writer, index=False, sheet_name="Meta")

    output_buffer.seek(0)
    st.download_button(
        label="📥 Download Excel Report",
        data=output_buffer.getvalue(),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info(
        "💡 Unggah file lalu klik **▶️ Proses**. "
        "Setelah diproses, filter SO tidak akan memproses ulang — "
        "hanya menampilkan data dari hasil yang sudah tersimpan."
    )
