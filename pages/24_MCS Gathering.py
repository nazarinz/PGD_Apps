import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO

from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="MCS Recap Generator", layout="wide")

# =====================================================
# 1️⃣ PREPARE FUNCTIONS
# =====================================================

def prepare_mcs(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()

    df = df.rename(columns={
        "Article No": "Article/No",
        "1ST Ex-factory date": "1st Ex-Factory Date",
        "Gender": "Age Group/Gender",
        "Bottom Tooling No.": "Bottom Tooling No",
        "Upper Tooling No.": "Upper Tooling No",
    })

    if "Size" not in df.columns:
        df["Size"] = None

    if "Factory Priority in Origo System" not in df.columns:
        df["Factory Priority in Origo System"] = None

    return df


def prepare_original(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()

    df = df.rename(columns={
        "Article": "Article/No",
        "Age Group": "Age Group/Gender",
    })

    if "Size" not in df.columns:
        df["Size"] = None

    if "Factory Priority in Origo System" not in df.columns:
        df["Factory Priority in Origo System"] = None

    return df


# =====================================================
# 2️⃣ FINAL COLUMN STRUCTURE (OUTPUT)
# =====================================================

final_cols = [
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
    "Size"
]

# =====================================================
# 6️⃣ SOCCER DETECTION (UPDATED)
# =====================================================

soccer_keywords = [
    "PREDATOR",
    "COPA",
    "CRAZYFAST",
    "F50",
    "MESSI",
    "EDGE",
    "ACCURACY",
    "FREAK",
    "SALA",
    "ADIPURE",
    "SAMBA TEAM",
    "SUPERSTAR JUDE",
]
soccer_pattern = "|".join(re.escape(k) for k in soccer_keywords)

# =====================================================
# 7️⃣ NORMALIZE AGE GROUP / GENDER
# =====================================================

normalize_mapping = {
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

# =====================================================
# 8️⃣ SIZE GENERATION (DON'T OVERWRITE)
# =====================================================

size_mapping = {
    "MEN": "8-",
    "WOMEN": "5-",
    "JUNIOR": "3",
    "CHILDREN": "11-K",
    "INFANT": "5-K",
    "KIDS": "11-K",
}

# =====================================================
# 9️⃣ SIZE RUN NORMALIZATION
# =====================================================

def clean_token(token: str) -> str:
    token = token.strip().upper()
    token = token.replace(" ", "")
    return token

def expand_range_numeric(start, end, is_kids=False):
    sizes = []
    start = int(float(start))
    end = int(float(end))
    for i in range(start, end + 1):
        if is_kids:
            sizes.append(f"{i}K")
            sizes.append(f"{i}-K")
        else:
            sizes.append(f"{i}")
            sizes.append(f"{i}-")
    return sizes

def normalize_size_run(size_run):
    if pd.isna(size_run):
        return None

    raw = str(size_run).upper()
    raw = raw.replace(";", ",")
    raw = raw.replace(" ", "")

    tokens = re.split(r",", raw)
    final_sizes = []

    for token in tokens:
        token = clean_token(token)
        if not token:
            continue

        kids_match = re.match(r"(\d+\.?\d*)K-(\d+\.?\d*)K", token)
        if kids_match:
            start, end = kids_match.groups()
            final_sizes.extend(expand_range_numeric(start, end, is_kids=True))
            continue

        adult_match = re.match(r"(\d+\.?\d*)-(\d+\.?\d*)", token)
        if adult_match and "K" not in token:
            start, end = adult_match.groups()
            final_sizes.extend(expand_range_numeric(start, end, is_kids=False))
            continue

        single_k_match = re.match(r"(\d+)K", token)
        if single_k_match:
            num = single_k_match.group(1)
            final_sizes.append(f"{num}K")
            final_sizes.append(f"{num}-K")
            continue

        single_match = re.match(r"(\d+)", token)
        if single_match:
            num = single_match.group(1)
            final_sizes.append(f"{num}")
            final_sizes.append(f"{num}-")
            continue

    seen = set()
    normalized = []
    for s in final_sizes:
        if s not in seen:
            seen.add(s)
            normalized.append(s)

    return ",".join(normalized)

# =====================================================
# EXCEL EXPORT (STYLED) TO BYTES
# =====================================================

def export_styled_excel(df: pd.DataFrame, sheet_name="MCS Recap") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]

        font_style = Font(name="Calibri", size=9)
        alignment_style = Alignment(horizontal="center", vertical="center")

        for row in ws.iter_rows():
            for cell in row:
                cell.font = font_style
                cell.alignment = alignment_style

        for col_cells in ws.columns:
            max_length = 0
            col_letter = get_column_letter(col_cells[0].column)
            for cell in col_cells:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max_length + 2

        for row in ws.iter_rows():
            ws.row_dimensions[row[0].row].height = 15

    return output.getvalue()

# =====================================================
# FILE LOADING HELPERS
# =====================================================

def list_sheets(uploaded_file) -> list[str]:
    xls = pd.ExcelFile(uploaded_file)
    return xls.sheet_names

def read_excel(uploaded_file, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(uploaded_file, sheet_name=sheet_name)

# =====================================================
# UI
# =====================================================

st.title("MCS Recap Generator (Deploy-ready)")

with st.sidebar:
    st.header("Upload Files")
    fw26_file = st.file_uploader("MCS FW26 (Excel)", type=["xlsx", "xls"], key="fw26")
    fw27_file = st.file_uploader("MCS FW27 (Excel)", type=["xlsx", "xls"], key="fw27")
    ss27_file = st.file_uploader("MCS SS27 (Excel)", type=["xlsx", "xls"], key="ss27")
    football_file = st.file_uploader("Original Football (Excel)", type=["xlsx", "xls"], key="football")

    st.divider()
    st.header("Options")
    preview_rows = st.number_input("Preview rows", min_value=5, max_value=200, value=20, step=5)

# Sheet selectors
def sheet_selector(label: str, f):
    if not f:
        st.info(f"Upload file untuk {label} dulu.")
        return None
    sheets = list_sheets(f)
    if len(sheets) == 1:
        return sheets[0]
    return st.selectbox(f"Sheet untuk {label}", sheets, index=0, key=f"sheet_{label}")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Sheet Selection")
    fw26_sheet = sheet_selector("FW26", fw26_file)
    fw27_sheet = sheet_selector("FW27", fw27_file)
    ss27_sheet = sheet_selector("SS27", ss27_file)
    football_sheet = sheet_selector("Football", football_file)

with col2:
    st.subheader("Status")
    missing = [name for name, f in [
        ("FW26", fw26_file), ("FW27", fw27_file), ("SS27", ss27_file), ("Football", football_file)
    ] if f is None]
    if missing:
        st.warning("Belum lengkap: " + ", ".join(missing))
    else:
        st.success("Semua file sudah diupload ✅")

run = st.button("Generate MCS Recap", type="primary", disabled=bool(missing))

if run:
    try:
        # Read
        df_mcsfw26 = read_excel(fw26_file, fw26_sheet)
        df_mcsfw27 = read_excel(fw27_file, fw27_sheet)
        df_mcsss27 = read_excel(ss27_file, ss27_sheet)
        df_originalfootball = read_excel(football_file, football_sheet)

        # Prepare + reindex
        df_fw26 = prepare_mcs(df_mcsfw26).reindex(columns=final_cols)
        df_fw27 = prepare_mcs(df_mcsfw27).reindex(columns=final_cols)
        df_ss27 = prepare_mcs(df_mcsss27).reindex(columns=final_cols)
        df_football = prepare_original(df_originalfootball).reindex(columns=final_cols)

        # Standardize types
        for df in [df_fw26, df_fw27, df_ss27, df_football]:
            df["Bottom Tooling No"] = pd.to_numeric(df["Bottom Tooling No"], errors="coerce")
            df["Upper Tooling No"] = pd.to_numeric(df["Upper Tooling No"], errors="coerce")
            df["LT"] = pd.to_numeric(df["LT"], errors="coerce")
            df["1st Ex-Factory Date"] = pd.to_datetime(df["1st Ex-Factory Date"], errors="coerce")

            df["Age Group/Gender"] = (
                df["Age Group/Gender"]
                .fillna("")
                .astype(str)
                .str.upper()
                .str.strip()
            )

            df["Size"] = df["Size"].replace(r"^\s*$", pd.NA, regex=True)

        # Concat
        df_all = pd.concat([df_fw26, df_fw27, df_ss27, df_football], ignore_index=True)

        # Soccer detection
        cat_soccer = (
            df_all["Category"]
            .fillna("")
            .astype(str)
            .str.upper()
            .str.contains(r"FOOTBALL\s*/\s*SOCCER", regex=True, na=False)
        )

        name_soccer = (
            df_all["Model Name"]
            .fillna("")
            .astype(str)
            .str.upper()
            .str.contains(soccer_pattern, na=False, regex=True)
        )

        df_all["Is_Soccer"] = cat_soccer | name_soccer

        # Normalize age group/gender
        df_all["Age Group/Gender Normalized"] = df_all["Age Group/Gender"].map(normalize_mapping)

        # Size generation (fill blanks only)
        generated_size = df_all["Age Group/Gender Normalized"].map(size_mapping)

        mask_special = (
            (df_all["Age Group/Gender Normalized"] == "KIDS") &
            (df_all["Is_Soccer"] == True)
        )
        generated_size.loc[mask_special] = "3"

        existing_size = df_all["Size"].replace(r"^\s*$", pd.NA, regex=True)
        df_all["Size"] = existing_size.combine_first(generated_size)

        # Size run normalized
        df_all["Size Run Normalized"] = df_all["Size Run"].apply(normalize_size_run)

        # Output cols
        output_cols = final_cols + ["Is_Soccer", "Age Group/Gender Normalized", "Size Run Normalized"]
        df_all = df_all.reindex(columns=output_cols)

        st.subheader("Preview Output")
        st.dataframe(df_all.head(int(preview_rows)), use_container_width=True)

        # Export
        today_str = datetime.today().strftime("%Y%m%d")
        file_name = f"MCS Recap - {today_str}.xlsx"
        excel_bytes = export_styled_excel(df_all, sheet_name="MCS Recap")

        st.download_button(
            label="Download Excel (Styled)",
            data=excel_bytes,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.success(f"✅ Generated: {file_name}")

    except Exception as e:
        st.exception(e)
