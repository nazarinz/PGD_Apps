"""
Arrival Report Comparison Tool
================================
Upload satu file Excel yang berisi dua data ASN No berbeda.
App akan otomatis memisahkan dua group tersebut dan membandingkannya
baris per baris berdasarkan INV Item.

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

  .hero {
    background: linear-gradient(135deg, #0f1923 0%, #0d1a0d 100%);
    border: 1px solid #2a3f2a;
    border-radius: 12px;
    padding: 28px 36px;
    margin-bottom: 24px;
  }
  .hero h1 {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.85rem;
    font-weight: 600;
    color: #6ee87a;
    margin: 0 0 6px 0;
    letter-spacing: -0.5px;
  }
  .hero p { color: #8b9ab0; margin: 0; font-size: 0.92rem; }

  .stat-row { display: flex; gap: 12px; flex-wrap: wrap; margin: 18px 0; }
  .stat-chip {
    background: #131920;
    border: 1px solid #222e3a;
    border-radius: 8px;
    padding: 12px 20px;
    text-align: center;
    min-width: 120px;
  }
  .stat-num {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.6rem;
    font-weight: 600;
    display: block;
  }
  .stat-lbl { font-size: 0.75rem; color: #8b9ab0; margin-top: 2px; display: block; }
  .chip-ok    .stat-num { color: #6ee87a; }
  .chip-diff  .stat-num { color: #f0854a; }
  .chip-total .stat-num { color: #5bc0de; }
  .chip-only  .stat-num { color: #e8c84a; }

  .asn-badges { display: flex; gap: 14px; margin: 14px 0 20px; flex-wrap: wrap; }
  .asn-badge {
    border-radius: 8px;
    padding: 10px 18px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.85rem;
    font-weight: 600;
  }
  .asn-a { background: #1a2a4a; color: #5bc0de; border: 1px solid #2a4a8a; }
  .asn-b { background: #2a1a4a; color: #c07af0; border: 1px solid #5a2a9a; }

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

  div[data-testid="stSidebar"] {
    background: #0d1117;
    border-right: 1px solid #1e2d3a;
  }
</style>
""", unsafe_allow_html=True)

# ─── Constants ───────────────────────────────────────────────────────────────
HEADER_COLS = [
    "Invoice No", "BL AWB No", "Invoice Date",
    "Factory No", "ETD Date", "Supplier", "Shiping Type",
    "Arrival Type", "Goods Kind", "BC Date", "BC No", "BC Type", "AJU No",
]

ITEM_COLS = [
    "INV Item", "ASN Item", "PO", "PO Item", "Material Code",
    "Invoice Desc", "Unit", "Qty", "Price", "Amount",
    "FOC", "Net Weight", "Gross Weight",
]

# mode: "code" = TRUE/FALSE, "text" = YES/NO, "num" = selisih angka
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

JOIN_KEY = "INV Item"


# ─── Core functions ──────────────────────────────────────────────────────────

def find_header_row(df_raw: pd.DataFrame) -> int:
    for i, row in df_raw.iterrows():
        vals = [str(v).strip() for v in row]
        if "INV Item" in vals and "ASN Item" in vals:
            return i
    return 4


def read_and_split(file, asn_col: str = "ASN No"):
    df_raw  = pd.read_excel(file, sheet_name=0, header=None)
    hdr_row = find_header_row(df_raw)
    cols    = [str(c).strip() for c in df_raw.iloc[hdr_row].tolist()]

    df = df_raw.iloc[hdr_row + 1:].copy()
    df.columns = cols
    df = df.dropna(how="all").reset_index(drop=True)

    for col in ["Qty", "Price", "Amount", "Net Weight", "Gross Weight"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    for col in df.select_dtypes("object").columns:
        df[col] = df[col].astype(str).str.strip()

    if asn_col not in df.columns:
        raise ValueError(f"Kolom '{asn_col}' tidak ditemukan. Kolom yang ada: {df.columns.tolist()}")

    unique_asns = list(dict.fromkeys(df[asn_col].tolist()))  # preserve order
    if len(unique_asns) < 2:
        return df, unique_asns[0] if unique_asns else "", pd.DataFrame(), "", pd.DataFrame()

    asn_a, asn_b = unique_asns[0], unique_asns[1]
    df_a = df[df[asn_col] == asn_a].copy().reset_index(drop=True)
    df_b = df[df[asn_col] == asn_b].copy().reset_index(drop=True)
    return df, asn_a, df_a, asn_b, df_b


def compare_value(val_a: str, val_b: str, mode: str):
    if mode == "code":
        return str(val_a == val_b).upper()
    elif mode == "text":
        return "YES" if val_a == val_b else "NO"
    elif mode == "num":
        try:
            return round(float(val_a) - float(val_b), 4)
        except (ValueError, TypeError):
            return "N/A"
    return "?"


def build_comparison(df_a: pd.DataFrame, df_b: pd.DataFrame,
                     asn_a: str, asn_b: str) -> pd.DataFrame:
    if JOIN_KEY not in df_a.columns or JOIN_KEY not in df_b.columns:
        st.error(f"Kolom '{JOIN_KEY}' tidak ditemukan.")
        return pd.DataFrame()

    da = df_a.copy(); da[JOIN_KEY] = da[JOIN_KEY].astype(str)
    db = df_b.copy(); db[JOIN_KEY] = db[JOIN_KEY].astype(str)

    merged = da.merge(db, on=JOIN_KEY, how="outer",
                      suffixes=("_A", "_B"), indicator=True)

    rows = []
    for _, row in merged.iterrows():
        out = {}
        out["ASN No (A)"] = asn_a
        out["ASN No (B)"] = asn_b

        for col in HEADER_COLS:
            col_a = col + "_A" if col + "_A" in row.index else col
            col_b = col + "_B" if col + "_B" in row.index else col
            out[col] = row.get(col_a, row.get(col_b, ""))

        out[JOIN_KEY] = row[JOIN_KEY]

        for col in ITEM_COLS:
            if col == JOIN_KEY:
                continue
            key = col + "_A" if col + "_A" in row.index else col
            out[f"{col} [A]"] = row.get(key, "")

        for col in ITEM_COLS:
            if col == JOIN_KEY:
                continue
            key = col + "_B" if col + "_B" in row.index else col
            out[f"{col} [B]"] = row.get(key, "")

        for col, mode in COMPARE_COLS.items():
            key_a = col + "_A" if col + "_A" in row.index else col
            key_b = col + "_B" if col + "_B" in row.index else col
            out[f"{col} Check"] = compare_value(
                str(row.get(key_a, "")), str(row.get(key_b, "")), mode
            )

        out["_merge"] = row["_merge"]
        rows.append(out)

    return pd.DataFrame(rows)


def style_checks(df: pd.DataFrame, check_cols: list):
    def cell_color(val):
        if val in ("YES", "TRUE") or val == 0 or val == 0.0:
            return "color:#6ee87a; font-weight:600"
        elif val in ("NO", "FALSE"):
            return "color:#f0854a; font-weight:600"
        elif isinstance(val, (int, float)) and val != 0:
            return "color:#f0d84a; font-weight:600"
        return ""
    present = [c for c in check_cols if c in df.columns]
    return df.style.applymap(cell_color, subset=present) if present else df.style


def row_has_diff(row, check_cols):
    for c in check_cols:
        v = row.get(c)
        if v in ("NO", "FALSE"):
            return True
        if isinstance(v, (int, float)) and v != 0:
            return True
    return False


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df_export = df.drop(columns=["_merge"], errors="ignore")
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df_export.to_excel(writer, sheet_name="Comparison", index=False)
        wb = writer.book
        ws = writer.sheets["Comparison"]

        check_col_names = [f"{c} Check" for c in COMPARE_COLS]

        hdr_fmt = wb.add_format({
            "bold": True, "bg_color": "#1a2a3a", "font_color": "#a0d8ef",
            "border": 1, "font_name": "Calibri", "font_size": 9,
            "text_wrap": True, "valign": "vcenter",
        })
        chk_hdr = wb.add_format({
            "bold": True, "bg_color": "#1a2a1a", "font_color": "#6ee87a",
            "border": 1, "font_name": "Calibri", "font_size": 9,
            "text_wrap": True, "valign": "vcenter",
        })
        ok_fmt   = wb.add_format({"font_color": "#27ae60", "bold": True})
        bad_fmt  = wb.add_format({"font_color": "#e74c3c", "bold": True})
        diff_fmt = wb.add_format({"font_color": "#f39c12", "bold": True})

        for col_num, col_name in enumerate(df_export.columns):
            ws.write(0, col_num, col_name, chk_hdr if col_name in check_col_names else hdr_fmt)
            if col_name in check_col_names:
                ws.set_column(col_num, col_num, 18)
            elif col_name in HEADER_COLS:
                ws.set_column(col_num, col_num, 14)
            else:
                ws.set_column(col_num, col_num, 13)

        n = len(df_export)
        for col_name in check_col_names:
            if col_name not in df_export.columns:
                continue
            idx = df_export.columns.tolist().index(col_name)
            for criteria, value, fmt in [
                ("==", '"YES"',   ok_fmt),
                ("==", '"TRUE"',  ok_fmt),
                ("==", '"NO"',    bad_fmt),
                ("==", '"FALSE"', bad_fmt),
                ("==", 0,         ok_fmt),
                ("<>", 0,         diff_fmt),
            ]:
                ws.conditional_format(1, idx, n, idx, {
                    "type": "cell", "criteria": criteria,
                    "value": value, "format": fmt,
                })
        ws.freeze_panes(1, 0)
    return buf.getvalue()


# ─── Sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Pengaturan")
    st.markdown("---")

    st.markdown("**Kolom ASN No**")
    asn_col_input = st.text_input(
        "Nama kolom yang berisi ASN No",
        value="ASN No", label_visibility="collapsed"
    )

    st.markdown("**Join Key (kunci match)**")
    join_key_input = st.text_input(
        "Kolom untuk match baris antar dua group",
        value="INV Item", label_visibility="collapsed"
    )
    if join_key_input.strip():
        JOIN_KEY = join_key_input.strip()

    st.markdown("---")
    st.markdown("""
    **Cara Pakai**
    1. Upload **satu file Excel** berisi dua group ASN berbeda
    2. App otomatis split dua group
    3. Baris di-match pakai **INV Item**
    4. Review hasil di tabel
    5. Download sebagai Excel
    """)
    st.markdown("---")
    st.caption("Arrival Report Compare v2.0")


# ─── Hero ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <h1>📦 Arrival Report Compare</h1>
  <p>Upload <strong>satu file Excel</strong> yang berisi dua data ASN No berbeda.
     App otomatis memisahkan dan membandingkan kedua group baris per baris.</p>
</div>
""", unsafe_allow_html=True)


# ─── Upload ──────────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Upload file Excel Arrival Report (satu file, dua ASN No)",
    type=["xlsx", "xls"],
)


# ─── Main ────────────────────────────────────────────────────────────────────
if uploaded:
    asn_col = asn_col_input.strip() if asn_col_input.strip() else "ASN No"

    with st.spinner("Membaca file…"):
        try:
            df_full, asn_a, df_a, asn_b, df_b = read_and_split(uploaded, asn_col)
        except Exception as e:
            st.error(f"Error membaca file: {e}")
            st.stop()

    if df_a.empty or df_b.empty:
        st.warning("Hanya ditemukan satu nilai ASN No. Pastikan file berisi dua group ASN berbeda.")
        if asn_col in df_full.columns:
            st.write("Nilai ASN No yang ditemukan:", df_full[asn_col].unique().tolist())
        st.stop()

    # Tampilkan info group
    st.markdown(f"""
    <div class="asn-badges">
      <div class="asn-badge asn-a">◈ Group A &nbsp;·&nbsp; {asn_a} &nbsp;·&nbsp; {len(df_a)} baris</div>
      <div class="asn-badge asn-b">◈ Group B &nbsp;·&nbsp; {asn_b} &nbsp;·&nbsp; {len(df_b)} baris</div>
    </div>
    """, unsafe_allow_html=True)

    with st.spinner("Membandingkan data…"):
        result = build_comparison(df_a, df_b, asn_a, asn_b)

    if result.empty:
        st.stop()

    check_cols = [f"{c} Check" for c in COMPARE_COLS if f"{c} Check" in result.columns]

    total_matched = (result["_merge"] == "both").sum()
    only_a_n      = (result["_merge"] == "left_only").sum()
    only_b_n      = (result["_merge"] == "right_only").sum()
    matched_rows  = result[result["_merge"] == "both"]
    has_diff_n    = int(matched_rows.apply(lambda r: row_has_diff(r, check_cols), axis=1).sum())
    fully_ok_n    = int(total_matched) - has_diff_n

    st.markdown(f"""
    <div class="stat-row">
      <div class="stat-chip chip-total">
        <span class="stat-num">{total_matched}</span>
        <span class="stat-lbl">Rows Matched</span>
      </div>
      <div class="stat-chip chip-ok">
        <span class="stat-num">{fully_ok_n}</span>
        <span class="stat-lbl">Fully Match</span>
      </div>
      <div class="stat-chip chip-diff">
        <span class="stat-num">{has_diff_n}</span>
        <span class="stat-lbl">Ada Perbedaan</span>
      </div>
      <div class="stat-chip chip-only">
        <span class="stat-num">{only_a_n}</span>
        <span class="stat-lbl">Hanya di A</span>
      </div>
      <div class="stat-chip chip-only">
        <span class="stat-num">{only_b_n}</span>
        <span class="stat-lbl">Hanya di B</span>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # Build display column order
    display_cols = (
        ["ASN No (A)", "ASN No (B)"]
        + [c for c in HEADER_COLS if c in result.columns]
        + [JOIN_KEY]
        + [f"{c} [A]" for c in ITEM_COLS if c != JOIN_KEY and f"{c} [A]" in result.columns]
        + [f"{c} [B]" for c in ITEM_COLS if c != JOIN_KEY and f"{c} [B]" in result.columns]
        + check_cols
    )

    def show_df(df_sub):
        available = [c for c in display_cols if c in df_sub.columns]
        df_show = df_sub[available].reset_index(drop=True)
        st.dataframe(style_checks(df_show, check_cols), use_container_width=True, height=500)

    tab_all, tab_diff, tab_a_only, tab_b_only = st.tabs([
        f"🗂️ Semua ({len(result)})",
        f"⚠️ Perbedaan ({has_diff_n})",
        f"◀ Hanya A ({only_a_n})",
        f"▶ Hanya B ({only_b_n})",
    ])

    with tab_all:
        show_df(result)

    with tab_diff:
        if has_diff_n:
            diff_df = matched_rows[
                matched_rows.apply(lambda r: row_has_diff(r, check_cols), axis=1)
            ]
            show_df(diff_df)
        else:
            st.success("✅ Tidak ada perbedaan di semua baris yang matched!")

    with tab_a_only:
        df_sub = result[result["_merge"] == "left_only"]
        if not df_sub.empty:
            show_df(df_sub)
        else:
            st.info("Semua baris Group A punya pasangan di Group B.")

    with tab_b_only:
        df_sub = result[result["_merge"] == "right_only"]
        if not df_sub.empty:
            show_df(df_sub)
        else:
            st.info("Semua baris Group B punya pasangan di Group A.")

    # Summary per Check Column
    with st.expander("📊 Summary per Kolom Check", expanded=False):
        summary = []
        for col in check_cols:
            if col not in matched_rows.columns:
                continue
            vals = matched_rows[col]
            mode = COMPARE_COLS.get(col.replace(" Check", ""), "text")
            if mode == "num":
                n_ok   = int((vals == 0).sum())
                n_diff = int((vals != 0).sum())
                summary.append({"Kolom": col, "Mode": "Numerik", "✅ Sama": n_ok, "⚠️ Beda": n_diff,
                                 "% Match": f"{100*n_ok/len(vals):.1f}%" if len(vals) else "–"})
            else:
                ok_val = "TRUE" if mode == "code" else "YES"
                n_ok   = int((vals == ok_val).sum())
                n_diff = int((vals != ok_val).sum())
                summary.append({"Kolom": col, "Mode": "Kode" if mode == "code" else "Teks",
                                 "✅ Sama": n_ok, "⚠️ Beda": n_diff,
                                 "% Match": f"{100*n_ok/len(vals):.1f}%" if len(vals) else "–"})
        st.dataframe(pd.DataFrame(summary), use_container_width=True, hide_index=True)

    # Download
    st.markdown('<p class="section-head">Export Hasil</p>', unsafe_allow_html=True)
    excel_bytes = to_excel_bytes(result)
    st.download_button(
        label="⬇️  Download Hasil Comparison (Excel)",
        data=excel_bytes,
        file_name="arrival_report_comparison.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.markdown("""
    <div style="text-align:center; padding:70px 20px; color:#4a6a7a;">
      <div style="font-size:4rem; margin-bottom:16px;">📂</div>
      <div style="font-family:'IBM Plex Mono',monospace; font-size:1rem; color:#3a6a5a;">
        Upload satu file Excel di atas untuk mulai
      </div>
      <div style="font-size:0.85rem; margin-top:10px; color:#2a4a3a;">
        File harus berisi dua group ASN No berbeda dalam satu sheet
      </div>
    </div>
    """, unsafe_allow_html=True)
