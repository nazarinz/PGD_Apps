# app.py
import re
import io
from difflib import SequenceMatcher
import pandas as pd
import streamlit as st

# ----------------- Fuzzy core -----------------
def normalize(s: str) -> str:
    s = (s or "").upper()
    s = re.sub(r"[_/\\\-]+", " ", s)
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def token_sort(s: str) -> str:
    toks = normalize(s).split()
    toks.sort()
    return " ".join(toks)

def token_prefix(s: str, n: int = 5) -> str:
    toks = normalize(s).split()
    pref = [t[:n] for t in toks if t]
    pref.sort()
    return " ".join(pref)

def similarity(a: str, b: str) -> float:
    a1, b1 = normalize(a), normalize(b)
    a2, b2 = token_sort(a), token_sort(b)
    a3, b3 = token_prefix(a, 5), token_prefix(b, 5)

    s1 = SequenceMatcher(None, a1, b1).ratio()
    s2 = SequenceMatcher(None, a2, b2).ratio()
    s3 = SequenceMatcher(None, a3, b3).ratio()
    return 0.50 * s1 + 0.25 * s2 + 0.25 * s3

def best_match(query: str, master_lines: list[str]) -> tuple[str, float]:
    best_full = ""
    best_score = -1.0
    q = (query or "").strip()

    for full in master_lines:
        full = (full or "").strip()
        if not full:
            continue
        base = full.split("_")[0].strip()  # compare tanpa suffix
        sc = similarity(q, base)
        if sc > best_score:
            best_score = sc
            best_full = full
    return best_full, best_score

def status_from_score(score: float, thr_match: float, thr_review: float) -> str:
    if score >= thr_match:
        return "MATCH"
    if score >= thr_review:
        return "REVIEW"
    return "NO MATCH"

# ----------------- Paste parser -----------------
def read_excel_paste(tsv_text: str) -> pd.DataFrame:
    """
    Menerima paste dari Excel (umumnya TSV).
    Ekspektasi 2 kolom: Name_List dan Matched_Master.
    Bisa ada header atau tidak.
    """
    tsv_text = (tsv_text or "").strip("\n")
    if not tsv_text.strip():
        return pd.DataFrame(columns=["Name_List", "Matched_Master"])

    # read as TSV with no header first
    df = pd.read_csv(io.StringIO(tsv_text), sep="\t", header=None, dtype=str)

    # If user pasted many columns, take first 2
    if df.shape[1] >= 2:
        df = df.iloc[:, :2]
    else:
        # If only 1 column pasted, treat as Name_List only
        df["Matched_Master"] = ""
        df = df.iloc[:, :2]

    df.columns = ["Name_List", "Matched_Master"]
    df = df.fillna("")

    # Detect header row
    first_a = normalize(df.iloc[0, 0])
    first_b = normalize(df.iloc[0, 1])
    header_like = (
        first_a in {"NAME_LIST", "NAME", "INPUT", "INPUT NAME", "NAMA"} or
        first_b in {"MATCHED_MASTER", "MASTER", "MASTER LIST", "MATCHED"}
    )
    if header_like:
        df = df.iloc[1:].reset_index(drop=True)

    # trim
    df["Name_List"] = df["Name_List"].astype(str).str.strip()
    df["Matched_Master"] = df["Matched_Master"].astype(str).str.strip()

    return df

def to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Matching") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

# ----------------- Streamlit UI -----------------
st.set_page_config(page_title="Name Matching (Paste Excel → XLSX)", layout="wide")
st.title("Name Matching (Paste dari Excel → Output Excel)")

with st.sidebar:
    st.subheader("Threshold")
    thr_match = st.slider("MATCH >=", 0.50, 0.99, 0.85, 0.01)
    thr_review = st.slider("REVIEW >=", 0.30, 0.95, 0.70, 0.01)
    st.caption("Rule: score >= MATCH → MATCH, >= REVIEW → REVIEW, else NO MATCH")

st.markdown(
    """
**Cara pakai**
1) Di Excel, siapkan 2 kolom: **Name_List** dan **Matched_Master**  
2) Block kedua kolom itu → **Copy** → paste di box bawah  
"""
)

paste = st.text_area(
    "Paste di sini (2 kolom dari Excel):",
    height=260,
    placeholder="Name_List\tMatched_Master\nRIXY GILANG PRAYOG\tRIXY GILANG PRAYOGA_110asa\nSYARIF HIDAYATULLO\tSYARIF HIDAYA TULLOH_DWHPM"
)

df_in = read_excel_paste(paste)

st.write(f"Rows terbaca: **{len(df_in)}**")

# Build master list from Matched_Master column (unique, non-empty)
master_lines = [x for x in df_in["Matched_Master"].tolist() if str(x).strip()]
# If master column empty, fallback: treat Name_List as master (rare)
if not master_lines:
    master_lines = [x for x in df_in["Name_List"].tolist() if str(x).strip()]

# Dedup keep order
seen = set()
master_unique = []
for m in master_lines:
    if m not in seen:
        seen.add(m)
        master_unique.append(m)

run = st.button("Run Matching", type="primary", disabled=(len(df_in) == 0 or len(master_unique) == 0))

if run:
    out_rows = []
    for i, name in enumerate(df_in["Name_List"].tolist(), start=1):
        bm, sc = best_match(name, master_unique)
        stt = status_from_score(sc, thr_match, thr_review)
        matched = bm if stt != "NO MATCH" else ""
        out_rows.append([i, name, matched, round(sc, 3), stt])

    df_out = pd.DataFrame(out_rows, columns=["Row#", "Input Name", "Matched Master", "Similarity", "Status"])
    st.dataframe(df_out, use_container_width=True, hide_index=True)

    xlsx_bytes = to_xlsx_bytes(df_out, sheet_name="Matching")
    st.download_button(
        "Download Excel (.xlsx)",
        data=xlsx_bytes,
        file_name="name_matching_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
