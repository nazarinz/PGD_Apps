"""
MCS Recap Tool — Single File Streamlit App
Upload 4 source files → process → download styled Excel
"""

import re
import logging
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(message)s")
logger = logging.getLogger(__name__)

# =====================================================
# CONFIG / CONSTANTS
# =====================================================

MCS_RENAME = {
    "Article No": "Article/No",
    "1ST Ex-factory date": "1st Ex-Factory Date",
    "Gender": "Age Group/Gender",
    "Bottom Tooling No.": "Bottom Tooling No",
    "Upper Tooling No.": "Upper Tooling No",
}

ORIGINAL_RENAME = {
    "Article": "Article/No",
    "Age Group": "Age Group/Gender",
}

FINAL_COLS = [
    "Season",
    "Category",
    "Model Name",
    "Model No",
    "Article/No",
    "Status",
    "Factory Priority in Origo System",
    "Development Type",
    "Developer",
    "Bottom Tooling No",
    "Upper Tooling No",
    "Last",
    "1st Ex-Factory Date",
    "LT",
    "Age Group/Gender",
    "Size Run",
    "Size",
]

OUTPUT_COLS = FINAL_COLS + [
    "Is_Soccer",
    "Age Group/Gender Normalized",
    "Size Run Normalized",
]

SOCCER_KEYWORDS = [
    "PREDATOR", "COPA", "CRAZYFAST", "F50", "MESSI",
    "EDGE", "ACCURACY", "FREAK", "SALA", "ADIPURE",
    "SAMBA TEAM", "SUPERSTAR JUDE",
]

AGE_NORMALIZE_MAP = {
    "U-JUNIOR": "JUNIOR",
    "JUNIOR": "JUNIOR",
    "UNISEX": "UNISEX",
    "M": "MEN",
    "MALE": "MEN",
    "MEN": "MEN",
    "W": "WOMEN",
    "FEMALE": "WOMEN",
    "WOMEN": "WOMEN",
    "CHILDREN": "CHILDREN",
    "C": "CHILDREN",
    "U-INFANTS": "INFANT",
    "INFANT": "INFANT",
    "ADULT": "ADULT",
    "KIDS": "KIDS",
}

SIZE_MAP = {
    "MEN": "8-",
    "WOMEN": "5-",
    "JUNIOR": "3",
    "CHILDREN": "11-K",
    "INFANT": "5-K",
    "KIDS": "11-K",
}


# =====================================================
# PIPELINE FUNCTIONS
# =====================================================

def prepare_mcs(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    df = df.rename(columns=MCS_RENAME)
    for col in ("Size", "Factory Priority in Origo System"):
        if col not in df.columns:
            df[col] = None
    return df


def prepare_original(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    df = df.rename(columns=ORIGINAL_RENAME)
    for col in ("Size", "Factory Priority in Origo System"):
        if col not in df.columns:
            df[col] = None
    return df


def standardize_types(df: pd.DataFrame) -> pd.DataFrame:
    for col in ("Bottom Tooling No", "Upper Tooling No", "LT"):
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df["1st Ex-Factory Date"] = pd.to_datetime(df["1st Ex-Factory Date"], errors="coerce")
    df["Age Group/Gender"] = df["Age Group/Gender"].fillna("").astype(str).str.upper().str.strip()
    df["Size"] = df["Size"].replace(r"^\s*$", pd.NA, regex=True)
    return df


def detect_soccer(df: pd.DataFrame) -> pd.DataFrame:
    pattern = "|".join(re.escape(k) for k in SOCCER_KEYWORDS)
    cat_soccer = (
        df["Category"].fillna("").astype(str).str.upper()
        .str.contains(r"FOOTBALL\s*/\s*SOCCER", regex=True, na=False)
    )
    name_soccer = (
        df["Model Name"].fillna("").astype(str).str.upper()
        .str.contains(pattern, regex=True, na=False)
    )
    df["Is_Soccer"] = cat_soccer | name_soccer
    return df


def normalize_age_group(df: pd.DataFrame) -> pd.DataFrame:
    df["Age Group/Gender Normalized"] = df["Age Group/Gender"].map(AGE_NORMALIZE_MAP)
    return df


def generate_size(df: pd.DataFrame) -> pd.DataFrame:
    generated = df["Age Group/Gender Normalized"].map(SIZE_MAP)
    mask_kids_soccer = (df["Age Group/Gender Normalized"] == "KIDS") & (df["Is_Soccer"])
    generated.loc[mask_kids_soccer] = "3"
    existing = df["Size"].replace(r"^\s*$", pd.NA, regex=True)
    df["Size"] = existing.combine_first(generated)
    return df


def expand_range(start: str, end: str, is_kids: bool = False) -> list:
    sizes = []
    for i in range(int(float(start)), int(float(end)) + 1):
        sizes += ([f"{i}K", f"{i}-K"] if is_kids else [str(i), f"{i}-"])
    return sizes


def normalize_size_run(value) -> str | None:
    if pd.isna(value):
        return None
    raw = str(value).upper().replace(";", ",").replace(" ", "")
    final = []
    for token in raw.split(","):
        token = token.strip()
        if not token:
            continue
        if m := re.match(r"(\d+\.?\d*)K-(\d+\.?\d*)K", token):
            final.extend(expand_range(*m.groups(), is_kids=True))
        elif m := re.match(r"(\d+\.?\d*)-(\d+\.?\d*)", token):
            if "K" not in token:
                final.extend(expand_range(*m.groups(), is_kids=False))
        elif m := re.match(r"(\d+)K", token):
            n = m.group(1)
            final += [f"{n}K", f"{n}-K"]
        elif m := re.match(r"(\d+)", token):
            n = m.group(1)
            final += [str(n), f"{n}-"]
    seen: set = set()
    deduped = [s for s in final if not (s in seen or seen.add(s))]  # type: ignore
    return ",".join(deduped)


def run_pipeline(
    df_mcsfw26: pd.DataFrame,
    df_mcsfw27: pd.DataFrame,
    df_mcsss27: pd.DataFrame,
    df_originalfootball: pd.DataFrame,
) -> pd.DataFrame:
    frames = [
        prepare_mcs(df_mcsfw26).reindex(columns=FINAL_COLS),
        prepare_mcs(df_mcsfw27).reindex(columns=FINAL_COLS),
        prepare_mcs(df_mcsss27).reindex(columns=FINAL_COLS),
        prepare_original(df_originalfootball).reindex(columns=FINAL_COLS),
    ]
    for df in frames:
        standardize_types(df)

    df_all = pd.concat(frames, ignore_index=True)
    df_all = (
        df_all
        .pipe(detect_soccer)
        .pipe(normalize_age_group)
        .pipe(generate_size)
    )
    df_all["Size Run Normalized"] = df_all["Size Run"].apply(normalize_size_run)
    return df_all.reindex(columns=OUTPUT_COLS)


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="MCS Recap")
        ws = writer.sheets["MCS Recap"]

        font = Font(name="Calibri", size=9)
        align = Alignment(horizontal="center", vertical="center")

        for row in ws.iter_rows():
            for cell in row:
                cell.font = font
                cell.alignment = align

        for col_cells in ws.columns:
            letter = get_column_letter(col_cells[0].column)
            max_len = max(
                (len(str(c.value)) for c in col_cells if c.value is not None),
                default=10,
            )
            ws.column_dimensions[letter].width = max_len + 2

        for row in ws.iter_rows():
            ws.row_dimensions[row[0].row].height = 15

    return buffer.getvalue()


# =====================================================
# STREAMLIT UI
# =====================================================

st.set_page_config(
    page_title="MCS Recap Tool",
    page_icon="👟",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .main .block-container { padding-top: 2rem; }
    div[data-testid="metric-container"] {
        background: #f0f2f6; border-radius: 8px; padding: 0.5rem 1rem;
    }
</style>
""", unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def read_excel_cached(raw: bytes, filename: str) -> pd.DataFrame:
    try:
        return pd.read_excel(raw, dtype=str)
    except Exception as e:
        raise ValueError(f"Cannot read '{filename}': {e}") from e


def safe_read(uploaded) -> pd.DataFrame | None:
    if uploaded is None:
        return None
    try:
        return read_excel_cached(uploaded.read(), uploaded.name)
    except ValueError as e:
        st.error(str(e))
        return None


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("👟 MCS Recap Tool")
    st.caption("Upload 4 file sumber untuk generate laporan Excel konsolidasi.")
    st.divider()
    st.subheader("📂 Upload Files")

    uf_fw26     = st.file_uploader("MCS FW26",          type=["xlsx", "xls"], key="fw26")
    uf_fw27     = st.file_uploader("MCS FW27",          type=["xlsx", "xls"], key="fw27")
    uf_ss27     = st.file_uploader("MCS SS27",          type=["xlsx", "xls"], key="ss27")
    uf_football = st.file_uploader("Original Football", type=["xlsx", "xls"], key="football")

    st.divider()
    run_btn = st.button("⚙️ Generate Report", use_container_width=True, type="primary")


# ── Main ──────────────────────────────────────────────────────────────────────
st.title("MCS Recap Generator")
st.caption("Konsolidasi FW26 · FW27 · SS27 · Football → satu Excel terformat.")

all_uploaded = all([uf_fw26, uf_fw27, uf_ss27, uf_football])

col1, col2, col3, col4 = st.columns(4)
for col, label, uf in zip(
    [col1, col2, col3, col4],
    ["MCS FW26", "MCS FW27", "MCS SS27", "Football"],
    [uf_fw26, uf_fw27, uf_ss27, uf_football],
):
    with col:
        st.metric(label=label, value="✅" if uf else "⬜")

if not all_uploaded:
    st.info("👈 Upload semua 4 file di sidebar untuk mengaktifkan Generate Report.", icon="ℹ️")

if run_btn:
    if not all_uploaded:
        st.warning("Upload semua 4 file terlebih dahulu.", icon="⚠️")
        st.stop()

    with st.status("Memproses file…", expanded=True) as status:
        try:
            st.write("📖 Membaca file…")
            df_fw26     = safe_read(uf_fw26)
            df_fw27     = safe_read(uf_fw27)
            df_ss27     = safe_read(uf_ss27)
            df_football = safe_read(uf_football)

            if None in [df_fw26, df_fw27, df_ss27, df_football]:
                st.error("Satu atau lebih file gagal dibaca.")
                st.stop()

            st.write("⚙️ Menjalankan pipeline…")
            df_result = run_pipeline(df_fw26, df_fw27, df_ss27, df_football)

            st.write("📊 Membuat Excel…")
            excel_bytes = to_excel_bytes(df_result)

            status.update(label="✅ Laporan siap!", state="complete")

        except Exception as e:
            logger.exception("Pipeline failed")
            status.update(label="❌ Gagal", state="error")
            st.error(f"**Error:** {e}")
            st.stop()

    st.divider()
    st.subheader("📊 Preview")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Rows",       f"{len(df_result):,}")
    m2.metric("Soccer Articles",  f"{df_result['Is_Soccer'].sum():,}")
    m3.metric("Seasons",          df_result["Season"].nunique())
    m4.metric("Unique Models",    df_result["Model No"].nunique())

    tab_all, tab_soccer = st.tabs(["All Data", "Soccer Only"])
    with tab_all:
        st.dataframe(df_result, use_container_width=True, height=400)
    with tab_soccer:
        st.dataframe(df_result[df_result["Is_Soccer"]], use_container_width=True, height=400)

    st.divider()
    filename = f"MCS Recap - {datetime.today().strftime('%Y%m%d')}.xlsx"
    st.download_button(
        label="⬇️ Download Excel Report",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
