"""
Arrival Report Comparison Tool
================================
Compare two Excel Arrival Report files side-by-side with automatic
difference highlighting and a downloadable result.

Run with:
    streamlit run arrival_report_compare.py

Requirements:
    pip install streamlit pandas openpyxl xlsxwriter
"""

import io
import pandas as pd
import streamlit as st

# ─── Page config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Arrival Report Compare",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Custom CSS ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');

  html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }

  .main { background: #0d1117; }

  /* Hero banner */
  .hero {
    background: linear-gradient(135deg, #1a1f2e 0%, #0f1923 60%, #1a2a1a 100%);
    border: 1px solid #2a3a2a;
    border-radius: 12px;
    padding: 28px 36px;
    margin-bottom: 24px;
  }
  .hero h1 {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.9rem;
    font-weight: 600;
    color: #6ee87a;
    margin: 0 0 6px 0;
    letter-spacing: -0.5px;
  }
  .hero p { color: #8b9ab0; margin: 0; font-size: 0.92rem; }

  /* Upload cards */
  .upload-card {
    background: #131920;
    border: 1.5px dashed #2a3f2a;
    border-radius: 10px;
    padding: 20px;
    transition: border-color 0.2s;
  }
  .upload-card:hover { border-color: #4a8f4a; }
  .upload-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.78rem;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    padding: 4px 10px;
    border-radius: 4px;
    margin-bottom: 10px;
    display: inline-block;
  }
  .label-a { background: #1a3a4a; color: #5bc0de; }
  .label-b { background: #3a1a4a; color: #c05bd8; }

  /* Stat chips */
  .stat-row { display: flex; gap: 12px; flex-wrap: wrap; margin: 18px 0; }
  .stat-chip {
    background: #131920;
    border: 1px solid #222e3a;
    border-radius: 8px;
    padding: 10px 18px;
    text-align: center;
    min-width: 110px;
  }
  .stat-num {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.5rem;
    font-weight: 600;
    display: block;
  }
  .stat-lbl { font-size: 0.75rem; color: #8b9ab0; }
  .chip-ok  .stat-num { color: #6ee87a; }
  .chip-diff .stat-num { color: #f0854a; }
  .chip-total .stat-num { color: #5bc0de; }
  .chip-only .stat-num { color: #e8c84a; }

  /* Table tweaks */
  .stDataFrame { border-radius: 8px; overflow: hidden; }

  /* Section headers */
  .section-head {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.75rem;
    font-weight: 600;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #4a7a8a;
    border-bottom: 1px solid #1e2d3a;
    padding-bottom: 6px;
    margin: 24px 0 14px 0;
  }

  /* Download button */
  .stDownloadButton > button {
    background: linear-gradient(90deg, #1a4a2a, #2a6a3a) !important;
    color: #a0f0a8 !important;
    border: none !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-weight: 600 !important;
    letter-spacing: 0.05em !important;
    border-radius: 8px !important;
    padding: 10px 24px !important;
  }
  .stDownloadButton > button:hover {
    background: linear-gradient(90deg, #2a6a3a, #3a8a4a) !important;
  }

  /* Badge for match status */
  .badge-ok   { color: #6ee87a; font-weight: 600; }
  .badge-diff { color: #f0854a; font-weight: 600; }

  div[data-testid="stSidebar"] {
    background: #0d1117;
    border-right: 1px solid #1e2d3a;
  }
</style>
""", unsafe_allow_html=True)

# ─── Constants ────────────────────────────────────────────────────────────────
HEADER_ROW   = 4        # 0-indexed row where column headers live
DATA_START   = 5        # 0-indexed first data row

# Columns that form the "document header" (shared / not compared)
HEADER_COLS = [
    "ASN No", "Invoice No", "BL AWB No", "Invoice Date",
    "Factory No", "ETD Date", "Supplier", "Shiping Type",
    "Arrival Type", "Goods Kind", "BC Date", "BC No", "BC Type", "AJU No",
]

# Item-level columns present in both files
ITEM_COLS = [
    "INV Item", "ASN Item", "PO", "PO Item", "Material Code",
    "Invoice Desc", "Unit", "Qty", "Price", "Amount",
    "FOC", "Net Weight", "Gross Weight",
]

# Which item columns to compare, and how
# "text"    → YES / NO
# "num"     → numeric difference  (0 = identical)
# "code"    → TRUE / FALSE  (TRUE = same, FALSE = different) — inverted display
COMPARE_COLS = {
    "Material Code": "code",
    "Invoice Desc":  "text",
    "Unit":          "text",
    "Qty":           "num",
    "Price":         "num",
    "Amount":        "num",
    "FOC":           "text",
    "Net Weight":    "num",
    "Gross Weight":  "num",
}

# Join key (must be unique per document)
JOIN_KEY = "INV Item"


# ─── Helpers ─────────────────────────────────────────────────────────────────

def detect_header_row(df_raw):
    """
    Auto-detect which row contains the expected header.
    Falls back to HEADER_ROW if not found.
    """
    for i, row in df_raw.iterrows():
        vals = [str(v).strip() for v in row]
        if "INV Item" in vals and "ASN Item" in vals:
            return i
    return HEADER_ROW


def read_excel(file) -> pd.DataFrame:
    """
    Read an uploaded Excel file and return a clean DataFrame.
    Auto-detects header row and drops empty rows.
    """
    df_raw = pd.read_excel(file, sheet_name=0, header=None)
    hdr    = detect_header_row(df_raw)
    cols   = df_raw.iloc[hdr].tolist()

    df = df_raw.iloc[hdr + 1:].copy()
    df.columns = cols
    df = df.dropna(how="all").reset_index(drop=True)

    # Normalise column names (strip whitespace)
    df.columns = [str(c).strip() for c in df.columns]

    # Coerce numeric columns
    for col in ["Qty", "Price", "Amount", "Net Weight", "Gross Weight"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Trim string columns
    for col in df.select_dtypes("object").columns:
        df[col] = df[col].astype(str).str.strip()

    return df


def compare_value(val_a, val_b, mode):
    """Return a comparison result cell based on mode."""
    if mode == "code":
        return str(val_a == val_b).upper()   # TRUE / FALSE
    elif mode == "text":
        return "YES" if val_a == val_b else "NO"
    elif mode == "num":
        try:
            diff = float(val_b) - float(val_a)
            return round(diff, 4)
        except (ValueError, TypeError):
            return "N/A"
    return "?"


def build_comparison(df_a: pd.DataFrame, df_b: pd.DataFrame) -> pd.DataFrame:
    """
    Merge the two DataFrames on JOIN_KEY and produce the full comparison table.
    Rows only in A → appear with blanks on the B side (and vice-versa).
    """
    # Ensure JOIN_KEY exists
    for label, df in [("File A", df_a), ("File B", df_b)]:
        if JOIN_KEY not in df.columns:
            st.error(f"{label} does not contain column '{JOIN_KEY}'. "
                     "Please check your file format.")
            return pd.DataFrame()

    # Stringify join key for reliable matching
    df_a = df_a.copy(); df_a[JOIN_KEY] = df_a[JOIN_KEY].astype(str)
    df_b = df_b.copy(); df_b[JOIN_KEY] = df_b[JOIN_KEY].astype(str)

    # Merge – outer join so unmatched rows show up
    merged = df_a.merge(df_b, on=JOIN_KEY, how="outer",
                        suffixes=("_A", "_B"), indicator=True)

    result_rows = []

    for _, row in merged.iterrows():
        out = {}

        # ── Header / shared columns (take from A; fall back to B)
        for col in HEADER_COLS:
            col_a = col + "_A" if col + "_A" in row.index else col
            col_b = col + "_B" if col + "_B" in row.index else col
            val = row.get(col_a, row.get(col_b, ""))
            out[col] = val

        # ── JOIN KEY
        out[JOIN_KEY] = row[JOIN_KEY]

        # ── File A item columns (excluding JOIN_KEY, already set)
        for col in ITEM_COLS:
            if col == JOIN_KEY:
                continue
            col_a = col + "_A" if col + "_A" in row.index else col
            out[f"{col}_A"] = row.get(col_a, "")

        # ── File B item columns
        for col in ITEM_COLS:
            if col == JOIN_KEY:
                continue
            col_b = col + "_B" if col + "_B" in row.index else col
            out[f"{col}_B"] = row.get(col_b, "")

        # ── Check columns
        for col, mode in COMPARE_COLS.items():
            col_a = col + "_A" if col + "_A" in row.index else col
            col_b = col + "_B" if col + "_B" in row.index else col
            val_a = row.get(col_a, "")
            val_b = row.get(col_b, "")
            out[f"{col} Check"] = compare_value(str(val_a), str(val_b), mode)

        out["_merge"] = row["_merge"]   # for stats
        result_rows.append(out)

    return pd.DataFrame(result_rows)


def colour_check(val):
    """Pandas Styler cell colourer for Check columns."""
    if val == "YES" or val == "TRUE" or val == 0 or val == 0.0:
        return "color: #6ee87a; font-weight: 600;"
    elif val == "NO" or val == "FALSE":
        return "color: #f0854a; font-weight: 600;"
    elif isinstance(val, (int, float)) and val != 0:
        return "color: #f0d84a; font-weight: 600;"
    return ""


def colour_merge(val):
    if val == "left_only":
        return "background-color: #1a2a1a;"
    elif val == "right_only":
        return "background-color: #2a1a1a;"
    return ""


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Export the comparison DataFrame to an in-memory Excel file."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        # Drop the _merge helper column for export
        df_export = df.drop(columns=["_merge"], errors="ignore")
        df_export.to_excel(writer, sheet_name="Comparison", index=False)

        wb  = writer.book
        ws  = writer.sheets["Comparison"]

        # Header format
        hdr_fmt = wb.add_format({
            "bold": True, "bg_color": "#1a2a3a",
            "font_color": "#a0d8ef", "border": 1,
            "font_name": "Calibri", "font_size": 9,
            "text_wrap": True, "valign": "vcenter",
        })
        # Check-col header format
        chk_fmt = wb.add_format({
            "bold": True, "bg_color": "#1a2a1a",
            "font_color": "#6ee87a", "border": 1,
            "font_name": "Calibri", "font_size": 9,
            "text_wrap": True, "valign": "vcenter",
        })
        # Cell formats
        ok_fmt   = wb.add_format({"font_color": "#27ae60", "bold": True})
        bad_fmt  = wb.add_format({"font_color": "#e74c3c", "bold": True})
        diff_fmt = wb.add_format({"font_color": "#f39c12", "bold": True})

        # Write headers with custom format
        check_col_names = [f"{c} Check" for c in COMPARE_COLS]
        for col_num, col_name in enumerate(df_export.columns):
            fmt = chk_fmt if col_name in check_col_names else hdr_fmt
            ws.write(0, col_num, col_name, fmt)

        # Set column widths
        for col_num, col_name in enumerate(df_export.columns):
            if col_name in HEADER_COLS:
                ws.set_column(col_num, col_num, 14)
            elif col_name in check_col_names:
                ws.set_column(col_num, col_num, 16)
            else:
                ws.set_column(col_num, col_num, 13)

        # Conditional formatting on check columns
        for col_name in check_col_names:
            if col_name not in df_export.columns:
                continue
            idx = df_export.columns.tolist().index(col_name)
            col_letter = chr(ord('A') + idx) if idx < 26 else "A"   # simplified
            n = len(df_export)
            ws.conditional_format(1, idx, n, idx, {
                "type": "cell", "criteria": "==",
                "value": '"YES"', "format": ok_fmt,
            })
            ws.conditional_format(1, idx, n, idx, {
                "type": "cell", "criteria": "==",
                "value": '"TRUE"', "format": ok_fmt,
            })
            ws.conditional_format(1, idx, n, idx, {
                "type": "cell", "criteria": "==",
                "value": '"NO"', "format": bad_fmt,
            })
            ws.conditional_format(1, idx, n, idx, {
                "type": "cell", "criteria": "==",
                "value": '"FALSE"', "format": bad_fmt,
            })
            ws.conditional_format(1, idx, n, idx, {
                "type": "cell", "criteria": "==",
                "value": 0, "format": ok_fmt,
            })
            ws.conditional_format(1, idx, n, idx, {
                "type": "cell", "criteria": "!=",
                "value": 0, "format": diff_fmt,
            })

        ws.freeze_panes(1, 0)
    return buf.getvalue()


# ─── Sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Settings")
    st.markdown("---")
    st.markdown("**Join Key**")
    join_key_input = st.text_input(
        "Column used to match rows between files",
        value=JOIN_KEY, label_visibility="collapsed"
    )
    if join_key_input.strip():
        JOIN_KEY = join_key_input.strip()

    st.markdown("**Header Row (0-indexed)**")
    hdr_row_input = st.number_input(
        "Row number containing column names",
        min_value=0, max_value=20, value=HEADER_ROW,
        label_visibility="collapsed",
    )
    HEADER_ROW = int(hdr_row_input)

    st.markdown("---")
    st.markdown("""
    **How to use**
    1. Upload **File A** and **File B**
    2. Both must share the same column layout
    3. Rows are matched by the **Join Key** column
    4. Review differences in the table
    5. Download the result as Excel
    """)
    st.markdown("---")
    st.caption("Arrival Report Compare v1.0")


# ─── Hero ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <h1>📦 Arrival Report Compare</h1>
  <p>Upload two Arrival Report Excel files to compare them line-by-line.
     Differences are highlighted automatically and can be exported.</p>
</div>
""", unsafe_allow_html=True)


# ─── File uploaders ──────────────────────────────────────────────────────────
col_a, col_b = st.columns(2)

with col_a:
    st.markdown('<span class="upload-label label-a">◈ FILE A — Source / Reference</span>',
                unsafe_allow_html=True)
    file_a = st.file_uploader(
        "Upload File A", type=["xlsx", "xls"],
        key="file_a", label_visibility="collapsed"
    )

with col_b:
    st.markdown('<span class="upload-label label-b">◈ FILE B — Target / Comparison</span>',
                unsafe_allow_html=True)
    file_b = st.file_uploader(
        "Upload File B", type=["xlsx", "xls"],
        key="file_b", label_visibility="collapsed"
    )


# ─── Main Logic ──────────────────────────────────────────────────────────────
if file_a and file_b:
    with st.spinner("Reading and comparing files…"):
        try:
            df_a = read_excel(file_a)
            df_b = read_excel(file_b)
        except Exception as e:
            st.error(f"Error reading files: {e}")
            st.stop()

    if df_a.empty or df_b.empty:
        st.warning("One or both files appear to be empty after parsing.")
        st.stop()

    result = build_comparison(df_a, df_b)

    if result.empty:
        st.stop()

    # ── Stats ────────────────────────────────────────────────────────────────
    check_cols = [f"{c} Check" for c in COMPARE_COLS if f"{c} Check" in result.columns]

    def row_has_diff(row):
        for c in check_cols:
            v = row[c]
            if v in ("NO", "FALSE"):
                return True
            if isinstance(v, (int, float)) and v != 0:
                return True
        return False

    total_rows   = len(result[result["_merge"] == "both"])
    only_a       = (result["_merge"] == "left_only").sum()
    only_b       = (result["_merge"] == "right_only").sum()
    matched      = result[result["_merge"] == "both"]
    has_diff     = matched.apply(row_has_diff, axis=1).sum()
    fully_ok     = total_rows - has_diff

    st.markdown(f"""
    <div class="stat-row">
      <div class="stat-chip chip-total">
        <span class="stat-num">{total_rows}</span>
        <span class="stat-lbl">Matched Rows</span>
      </div>
      <div class="stat-chip chip-ok">
        <span class="stat-num">{fully_ok}</span>
        <span class="stat-lbl">Fully Match</span>
      </div>
      <div class="stat-chip chip-diff">
        <span class="stat-num">{has_diff}</span>
        <span class="stat-lbl">With Differences</span>
      </div>
      <div class="stat-chip chip-only">
        <span class="stat-num">{only_a}</span>
        <span class="stat-lbl">Only in A</span>
      </div>
      <div class="stat-chip chip-only">
        <span class="stat-num">{only_b}</span>
        <span class="stat-lbl">Only in B</span>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Tabs ─────────────────────────────────────────────────────────────────
    tab_all, tab_diff, tab_a_only, tab_b_only = st.tabs([
        "🗂️ All Rows",
        f"⚠️ Differences ({has_diff})",
        f"◀ Only in A ({only_a})",
        f"▶ Only in B ({only_b})",
    ])

    display_cols = (
        [c for c in HEADER_COLS if c in result.columns]
        + [JOIN_KEY]
        + [f"{c}_A" for c in ITEM_COLS if c != JOIN_KEY and f"{c}_A" in result.columns]
        + [f"{c}_B" for c in ITEM_COLS if c != JOIN_KEY and f"{c}_B" in result.columns]
        + [c for c in check_cols if c in result.columns]
    )

    # Rename display columns for readability
    rename_map = {}
    for c in ITEM_COLS:
        if c == JOIN_KEY:
            continue
        if f"{c}_A" in result.columns:
            rename_map[f"{c}_A"] = f"{c} [A]"
        if f"{c}_B" in result.columns:
            rename_map[f"{c}_B"] = f"{c} [B]"

    def show_table(df_sub, tab):
        df_show = df_sub[display_cols].rename(columns=rename_map)
        # Apply styling
        chk_renamed = [rename_map.get(c, c) for c in check_cols]
        styled = (
            df_show.style
            .applymap(colour_check, subset=[c for c in chk_renamed if c in df_show.columns])
        )
        tab.dataframe(styled, use_container_width=True, height=520)

    show_table(result, tab_all)

    if has_diff:
        diff_mask = result["_merge"] == "both"
        diff_rows = result[diff_mask & result.apply(row_has_diff, axis=1)]
        show_table(diff_rows, tab_diff)
    else:
        tab_diff.success("✅ No differences found between matched rows!")

    if only_a:
        show_table(result[result["_merge"] == "left_only"], tab_a_only)
    else:
        tab_a_only.info("No rows exclusive to File A.")

    if only_b:
        show_table(result[result["_merge"] == "right_only"], tab_b_only)
    else:
        tab_b_only.info("No rows exclusive to File B.")

    # ── Download ─────────────────────────────────────────────────────────────
    st.markdown('<p class="section-head">Export</p>', unsafe_allow_html=True)
    excel_bytes = to_excel_bytes(result)
    st.download_button(
        label="⬇️  Download Comparison as Excel",
        data=excel_bytes,
        file_name="arrival_report_comparison.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

elif file_a or file_b:
    st.info("Please upload **both** File A and File B to start the comparison.")

else:
    # Empty state
    st.markdown("""
    <div style="text-align:center; padding: 60px 20px; color: #4a6a7a;">
        <div style="font-size: 4rem; margin-bottom: 16px;">📂</div>
        <div style="font-family: 'IBM Plex Mono', monospace; font-size: 1rem; color: #3a5a6a;">
            Upload two Excel files above to begin
        </div>
        <div style="font-size: 0.85rem; margin-top: 8px; color: #2a4a5a;">
            Both files must use the standard Arrival Report column layout
        </div>
    </div>
    """, unsafe_allow_html=True)
