import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

st.set_page_config(page_title="Order vs Carton Form Comparator", page_icon="📦", layout="wide")

# ── Styling helpers ──────────────────────────────────────────────────────────
GREEN   = "00C853"
RED     = "D50000"
YELLOW  = "FFD600"
BLUE    = "1565C0"
LGRAY   = "F5F5F5"
DGRAY   = "424242"
WHITE   = "FFFFFF"
ORANGE  = "E65100"

def hdr_font(color=WHITE, bold=True, size=10):
    return Font(name="Arial", bold=bold, color=color, size=size)

def cell_font(bold=False, color="000000", size=10):
    return Font(name="Arial", bold=bold, color=color, size=size)

def fill(hex_color):
    return PatternFill("solid", start_color=hex_color, fgColor=hex_color)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def thin_border():
    s = Side(style="thin", color="BDBDBD")
    return Border(left=s, right=s, top=s, bottom=s)

def apply_cell(ws, row, col, value, bg=None, font=None, align=None, border=True, num_format=None):
    c = ws.cell(row=row, column=col, value=value)
    if bg:        c.fill = fill(bg)
    if font:      c.font = font
    if align:     c.alignment = align
    if border:    c.border = thin_border()
    if num_format: c.number_format = num_format
    return c

# ── Excel Parser ─────────────────────────────────────────────────────────────
def parse_excel(file) -> pd.DataFrame:
    df = pd.read_excel(file, dtype=str)
    df.columns = df.columns.str.strip()
    return df

def extract_excel_summary(df: pd.DataFrame) -> dict:
    r0 = df.iloc[0]
    summary = {
        "PO Number":        r0.get("Order #", ""),
        "Customer Order":   r0.get("Market PO Number", ""),
        "Article Number":   r0.get("Article Number", ""),
        "Model Name":       r0.get("Model Name", ""),
        "Ship Method":      r0.get("Shipment Method", ""),
        "Pack Mode":        r0.get("VAS/SHAS L15 – Packing Mode", ""),
        "Currency":         r0.get("Currency", ""),
        "Country":          r0.get("Country/Region", ""),
        "Final Destination":r0.get("FinalDestinationName", ""),
    }
    return summary

def extract_excel_sizes(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, r in df.iterrows():
        uk   = str(r.get("Manufacturing Size", "")).strip()
        us   = str(r.get("Customer Size", "")).strip()
        qty  = r.get("Quantity", "0")
        coa  = str(r.get("Contract Outline Agreement Number", "")).strip()
        line = str(r.get("Item Line Number", "")).strip()
        try:
            qty = int(float(qty))
        except Exception:
            qty = 0
        rows.append({"UK Size": uk, "US Size": us, "XL Qty": qty,
                     "Contract OA": coa, "Line": line})
    return pd.DataFrame(rows)

# ── PDF Parser ───────────────────────────────────────────────────────────────
def extract_text_from_pdf(file) -> str:
    text = ""
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text += (page.extract_text() or "") + "\n"
    return text

def parse_pdf_header(text: str) -> dict:
    def find(pattern, default=""):
        m = re.search(pattern, text)
        return m.group(1).strip() if m else default

    return {
        "PO Number":      find(r"PO NO[.:]?\s*([\w]+)"),
        "Customer Order": find(r"CUST\.O/N\s*([\w]+)"),
        "Article Number": find(r"ART\. NO[.:]?\s*([\w]+)"),
        "Model Name":     find(r"Model\s*[:\s]\s*([A-Z][A-Z0-9 ]+?)(?:\n|Arr|Ship|Coun|Sub|End|Sale|Cust|OUTER)"),
        "Ship Method":    find(r"ShipType[:\s]+[\d]+\s+([A-Za-z]+)"),
        "Total Pairs":    find(r"Pair[:\s]+([\d]+)\s+Pairs"),
        "Total Cartons":  find(r"Ctns[:\s]+([\d]+)\s+Ctn"),
        "Pack Mode":      find(r"(SSP|MSP|USSP|SP)\s+TOTAL"),
    }

def parse_pdf_sizes(text: str) -> pd.DataFrame:
    """
    Parse table rows like:
      1-2   2   6   12   4   5-   4.080  ...
      3-4   2   6   12   4-  6    4.200  ...
    """
    pattern = re.compile(
        r"(\d+[-–]\d+|\d+)\s+"     # ctn range
        r"(\d+)\s+"                 # num_ctns
        r"(\d+)\s+"                 # qty_per_ctn
        r"(\d+)\s+"                 # total_qty
        r"([\d]+[-]?)\s+"           # UK size
        r"([\d]+[-]?)\s+"           # US size
        r"[\d.]+"                   # nw per ctn (skip rest)
    )
    rows = []
    for m in pattern.finditer(text):
        ctn_range    = m.group(1)
        num_ctns     = int(m.group(2))
        qty_per_ctn  = int(m.group(3))
        total_qty    = int(m.group(4))
        uk_size      = m.group(5).strip()
        us_size      = m.group(6).strip()
        rows.append({
            "CTN Range":    ctn_range,
            "Num CTNs":     num_ctns,
            "Qty/CTN":      qty_per_ctn,
            "PDF Qty":      total_qty,
            "UK Size":      uk_size,
            "US Size":      us_size,
        })
    return pd.DataFrame(rows)

# ── Comparison Logic ─────────────────────────────────────────────────────────
def compare_headers(xl_hdr: dict, pdf_hdr: dict) -> pd.DataFrame:
    fields = [
        ("PO Number",      xl_hdr["PO Number"],        pdf_hdr["PO Number"]),
        ("Customer Order", xl_hdr["Customer Order"],    pdf_hdr["Customer Order"]),
        ("Article Number", xl_hdr["Article Number"],    pdf_hdr["Article Number"]),
        ("Model Name",     xl_hdr["Model Name"],        pdf_hdr["Model Name"].strip()),
        ("Ship Method",    xl_hdr["Ship Method"],       pdf_hdr["Ship Method"]),
        ("Pack Mode",      xl_hdr["Pack Mode"],         pdf_hdr["Pack Mode"]),
    ]
    rows = []
    for label, xv, pv in fields:
        match = xv.strip().upper() == pv.strip().upper()
        rows.append({"Field": label, "Excel Value": xv, "PDF Value": pv,
                     "Status": "MATCH" if match else "MISMATCH"})
    return pd.DataFrame(rows)

def compare_sizes(xl_sizes: pd.DataFrame, pdf_sizes: pd.DataFrame) -> pd.DataFrame:
    merged = pd.merge(
        xl_sizes[["UK Size","US Size","XL Qty","Contract OA","Line"]],
        pdf_sizes[["UK Size","US Size","PDF Qty","CTN Range","Num CTNs","Qty/CTN"]],
        on=["UK Size","US Size"], how="outer"
    )
    merged["XL Qty"]  = merged["XL Qty"].fillna(0).astype(int)
    merged["PDF Qty"] = merged["PDF Qty"].fillna(0).astype(int)
    merged["Diff"]    = merged["XL Qty"] - merged["PDF Qty"]

    def status(row):
        if pd.isna(row.get("CTN Range")) or row.get("CTN Range","") == "":
            return "ONLY IN EXCEL"
        if pd.isna(row.get("Line")) or str(row.get("Line","")).strip() == "":
            return "ONLY IN PDF"
        if row["Diff"] != 0:
            return "MISMATCH"
        if str(row.get("Contract OA","")).strip() in ("", "nan"):
            return "MATCH – COA MISSING"
        return "MATCH"

    merged["Status"] = merged.apply(status, axis=1)
    return merged

# ── Excel Output Builder ──────────────────────────────────────────────────────
def build_excel_output(xl_hdr, pdf_hdr, hdr_cmp, size_cmp, xl_sizes, pdf_sizes, xl_raw):
    wb = Workbook()

    # ── Sheet 1: Summary Dashboard ────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Summary"
    ws1.sheet_view.showGridLines = False

    # Title block
    ws1.merge_cells("A1:H1")
    c = ws1["A1"]
    c.value = "📦  ORDER vs CARTON FORM — COMPARISON REPORT"
    c.font  = Font(name="Arial", bold=True, size=14, color=WHITE)
    c.fill  = fill(DGRAY)
    c.alignment = center()
    ws1.row_dimensions[1].height = 32

    ws1.merge_cells("A2:H2")
    c = ws1["A2"]
    c.value = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  Article: {xl_hdr.get('Article Number','')}  |  PO: {xl_hdr.get('PO Number','')}"
    c.font  = Font(name="Arial", size=9, color="616161")
    c.fill  = fill(LGRAY)
    c.alignment = left()
    ws1.row_dimensions[2].height = 18

    # KPI boxes (row 4)
    total       = len(size_cmp)
    matches     = (size_cmp["Status"].str.startswith("MATCH")).sum()
    mismatches  = (size_cmp["Status"] == "MISMATCH").sum()
    coa_warn    = (size_cmp["Status"] == "MATCH – COA MISSING").sum()
    xl_total    = xl_sizes["XL Qty"].sum()
    pdf_total   = pdf_sizes["PDF Qty"].sum()

    kpis = [
        ("Total Sizes", total,     BLUE),
        ("Match",       matches,   GREEN),
        ("Mismatch",    mismatches,RED if mismatches>0 else GREEN),
        ("COA Warning", coa_warn,  YELLOW if coa_warn>0 else GREEN),
        ("Excel Pairs", xl_total,  DGRAY),
        ("PDF Pairs",   pdf_total, DGRAY if xl_total==pdf_total else RED),
    ]
    ws1.row_dimensions[4].height = 14
    ws1.row_dimensions[5].height = 36
    ws1.row_dimensions[6].height = 20
    for i, (lbl, val, col) in enumerate(kpis, start=1):
        ws1.cell(5, i).value     = val
        ws1.cell(5, i).font      = Font(name="Arial", bold=True, size=16, color=WHITE)
        ws1.cell(5, i).fill      = fill(col)
        ws1.cell(5, i).alignment = center()
        ws1.cell(6, i).value     = lbl
        ws1.cell(6, i).font      = Font(name="Arial", size=9, color="616161")
        ws1.cell(6, i).alignment = center()

    # Header comparison table
    ws1.row_dimensions[8].height = 16
    for col, hdr in enumerate(["Field", "Excel Value", "PDF Value", "Status"], start=1):
        apply_cell(ws1, 8, col, hdr, bg=DGRAY, font=hdr_font(), align=center())

    STATUS_COLOR = {"MATCH": GREEN, "MISMATCH": RED}
    for ri, row in hdr_cmp.iterrows():
        r = ri + 9
        ws1.row_dimensions[r].height = 16
        apply_cell(ws1, r, 1, row["Field"],       bg=LGRAY, font=cell_font(bold=True), align=left())
        apply_cell(ws1, r, 2, row["Excel Value"],              font=cell_font(), align=left())
        apply_cell(ws1, r, 3, row["PDF Value"],                font=cell_font(), align=left())
        sc = STATUS_COLOR.get(row["Status"], YELLOW)
        apply_cell(ws1, r, 4, row["Status"], bg=sc, font=hdr_font(color=WHITE if sc!=YELLOW else "000000"), align=center())

    # Column widths sheet1
    for col, w in zip("ABCDEFGH", [22, 28, 28, 16, 14, 14, 14, 14]):
        ws1.column_dimensions[col].width = w

    # ── Sheet 2: Size Comparison ──────────────────────────────────────────────
    ws2 = wb.create_sheet("Size Comparison")
    ws2.sheet_view.showGridLines = False

    ws2.merge_cells("A1:K1")
    c = ws2["A1"]
    c.value = "SIZE QTY COMPARISON — Excel vs Carton Form PDF"
    c.font  = Font(name="Arial", bold=True, size=12, color=WHITE)
    c.fill  = fill(DGRAY)
    c.alignment = center()
    ws2.row_dimensions[1].height = 26

    cols = ["UK Size","US Size","Excel Line","Excel Qty","PDF CTN Range",
            "PDF CTNs","Qty/CTN","PDF Qty","Difference","Contract OA","Status"]
    for ci, h in enumerate(cols, 1):
        apply_cell(ws2, 2, ci, h, bg="37474F", font=hdr_font(), align=center())
    ws2.row_dimensions[2].height = 18

    STATUS_ROW_BG = {
        "MATCH":             (WHITE,   "000000"),
        "MATCH – COA MISSING":(YELLOW, "000000"),
        "MISMATCH":          (RED,     WHITE),
        "ONLY IN EXCEL":     (ORANGE,  WHITE),
        "ONLY IN PDF":       (ORANGE,  WHITE),
    }

    for ri, row in size_cmp.iterrows():
        r  = ri + 3
        st = row["Status"]
        bg, fc = STATUS_ROW_BG.get(st, (WHITE, "000000"))
        ws2.row_dimensions[r].height = 16

        vals = [
            row.get("UK Size",""),   row.get("US Size",""),
            row.get("Line",""),      int(row["XL Qty"]),
            row.get("CTN Range",""), row.get("Num CTNs",""),
            row.get("Qty/CTN",""),   int(row["PDF Qty"]),
            int(row["Diff"]),        row.get("Contract OA",""),
            st,
        ]
        for ci, v in enumerate(vals, 1):
            bgi = bg if ci == 11 or bg != WHITE else WHITE
            fci = fc if ci == 11 else "000000"
            apply_cell(ws2, r, ci, v,
                       bg=bgi if ci==11 else (LGRAY if ri%2==0 else WHITE),
                       font=cell_font(bold=(ci==11), color=fci if ci==11 else "000000"),
                       align=center())
        # override status cell fully
        sc_bg, sc_fc = STATUS_ROW_BG.get(st, (WHITE,"000000"))
        apply_cell(ws2, r, 11, st, bg=sc_bg, font=cell_font(bold=True, color=sc_fc), align=center())

    # Totals row
    tr = len(size_cmp) + 3
    ws2.row_dimensions[tr].height = 18
    apply_cell(ws2, tr, 1, "TOTAL", bg=DGRAY, font=hdr_font(), align=center())
    apply_cell(ws2, tr, 4, int(xl_sizes["XL Qty"].sum()), bg=DGRAY, font=hdr_font(), align=center())
    apply_cell(ws2, tr, 8, int(pdf_sizes["PDF Qty"].sum()), bg=DGRAY, font=hdr_font(), align=center())
    diff_total = int(xl_sizes["XL Qty"].sum()) - int(pdf_sizes["PDF Qty"].sum())
    apply_cell(ws2, tr, 9, diff_total,
               bg=RED if diff_total!=0 else GREEN,
               font=hdr_font(), align=center())
    for ci in [2,3,5,6,7,10,11]:
        apply_cell(ws2, tr, ci, "", bg=DGRAY, font=hdr_font(), align=center())

    for col, w in zip("ABCDEFGHIJK", [10,10,10,12,16,10,10,12,14,20,22]):
        ws2.column_dimensions[col].width = w

    # ── Sheet 3: Discrepancies ────────────────────────────────────────────────
    ws3 = wb.create_sheet("Discrepancies")
    ws3.sheet_view.showGridLines = False

    ws3.merge_cells("A1:F1")
    c = ws3["A1"]
    c.value = "⚠  DISCREPANCY DETAIL"
    c.font  = fill_font = Font(name="Arial", bold=True, size=12, color=WHITE)
    c.fill  = fill(ORANGE)
    c.alignment = center()
    ws3.row_dimensions[1].height = 26

    issues = size_cmp[size_cmp["Status"] != "MATCH"].copy()

    if issues.empty:
        ws3["A3"] = "✓  No discrepancies found. All sizes match."
        ws3["A3"].font = Font(name="Arial", bold=True, color=GREEN, size=11)
    else:
        hdrs = ["UK Size","US Size","Excel Qty","PDF Qty","Difference","Issue"]
        for ci, h in enumerate(hdrs, 1):
            apply_cell(ws3, 2, ci, h, bg="37474F", font=hdr_font(), align=center())

        for ri, row in issues.iterrows():
            r = ri - issues.index[0] + 3
            ws3.row_dimensions[r].height = 16
            st = row["Status"]
            sc_bg, sc_fc = STATUS_ROW_BG.get(st, (YELLOW, "000000"))
            vals = [row.get("UK Size",""), row.get("US Size",""),
                    int(row["XL Qty"]), int(row["PDF Qty"]), int(row["Diff"]), st]
            for ci, v in enumerate(vals, 1):
                bg = sc_bg if ci==6 else (LGRAY if ri%2==0 else WHITE)
                fc = sc_fc if ci==6 else "000000"
                apply_cell(ws3, r, ci, v, bg=bg, font=cell_font(bold=(ci==6), color=fc), align=center())

    # COA missing note
    coa_rows = size_cmp[size_cmp["Status"] == "MATCH – COA MISSING"]
    note_row = len(issues) + 5
    if not coa_rows.empty:
        ws3.merge_cells(f"A{note_row}:F{note_row}")
        c = ws3[f"A{note_row}"]
        c.value = f"⚠  Contract OA Number missing on {len(coa_rows)} line(s): sizes " + \
                  ", ".join(coa_rows["UK Size"].tolist())
        c.font  = Font(name="Arial", bold=True, size=10, color="7B3F00")
        c.fill  = fill(YELLOW)
        c.alignment = left()

    for col, w in zip("ABCDEF", [10,10,14,14,14,26]):
        ws3.column_dimensions[col].width = w

    # ── Sheet 4: Raw Excel Data ───────────────────────────────────────────────
    ws4 = wb.create_sheet("Raw Excel Data")
    ws4.sheet_view.showGridLines = False

    for ci, col_name in enumerate(xl_raw.columns, 1):
        apply_cell(ws4, 1, ci, col_name, bg=DGRAY, font=hdr_font(size=9), align=center())

    for ri, row in xl_raw.iterrows():
        for ci, val in enumerate(row, 1):
            apply_cell(ws4, ri+2, ci, val,
                       font=Font(name="Arial", size=9),
                       align=left(),
                       bg=LGRAY if ri%2==0 else WHITE)

    for ci, col_name in enumerate(xl_raw.columns, 1):
        ws4.column_dimensions[get_column_letter(ci)].width = max(12, len(str(col_name))+2)

    ws4.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Streamlit UI ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { font-family: Arial, sans-serif; }
    .block-container { padding-top: 1.5rem; }
    .metric-card { background:#f5f5f5; border-radius:8px; padding:12px 16px; text-align:center; }
</style>
""", unsafe_allow_html=True)

st.title("📦 Order vs Carton Form Comparator")
st.caption("Upload file Excel (order lines) dan file PDF (carton form) untuk perbandingan otomatis.")

col1, col2 = st.columns(2)
with col1:
    st.subheader("📊 Excel — Order Lines")
    xl_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx","xls"])

with col2:
    st.subheader("📄 PDF — Carton Form")
    pdf_file = st.file_uploader("Upload file PDF Carton Form", type=["pdf"])

if xl_file and pdf_file:
    with st.spinner("Memproses data..."):
        # Parse
        xl_raw   = parse_excel(xl_file)
        xl_hdr   = extract_excel_summary(xl_raw)
        xl_sizes = extract_excel_sizes(xl_raw)

        pdf_text   = extract_text_from_pdf(pdf_file)
        pdf_hdr    = parse_pdf_header(pdf_text)
        pdf_sizes  = parse_pdf_sizes(pdf_text)

        # Compare
        hdr_cmp  = compare_headers(xl_hdr, pdf_hdr)
        size_cmp = compare_sizes(xl_sizes, pdf_sizes)

    # KPI row
    total      = len(size_cmp)
    match_cnt  = (size_cmp["Status"].str.startswith("MATCH")).sum()
    mismatch   = (size_cmp["Status"] == "MISMATCH").sum()
    warn_cnt   = (size_cmp["Status"] == "MATCH – COA MISSING").sum()
    xl_total   = xl_sizes["XL Qty"].sum()
    pdf_total  = pdf_sizes["PDF Qty"].sum()

    k1,k2,k3,k4,k5,k6 = st.columns(6)
    k1.metric("Total Sizes",  total)
    k2.metric("✅ Match",      match_cnt)
    k3.metric("❌ Mismatch",   mismatch,   delta_color="inverse")
    k4.metric("⚠️ COA Missing",warn_cnt,   delta_color="inverse")
    k5.metric("Excel Pairs",  int(xl_total))
    k6.metric("PDF Pairs",    int(pdf_total),
              delta=int(xl_total-pdf_total) if xl_total!=pdf_total else None,
              delta_color="inverse")

    st.divider()

    # Header comparison
    col_h, col_s = st.columns([1,2])
    with col_h:
        st.subheader("📋 Header Fields")
        def style_hdr(row):
            if row["Status"] == "MATCH":
                return ["","","","background-color:#c8f7c5"]
            return ["","","","background-color:#ffcdd2"]
        st.dataframe(hdr_cmp.style.apply(style_hdr, axis=1), use_container_width=True, hide_index=True)

    with col_s:
        st.subheader("📐 Size Qty Comparison")
        def style_size(row):
            s = row.get("Status","")
            if s == "MATCH":             bg = "#c8f7c5"
            elif s == "MATCH – COA MISSING": bg = "#fff9c4"
            elif s == "MISMATCH":        bg = "#ffcdd2"
            else:                        bg = "#ffe0cc"
            return [""] * (len(row)-1) + [f"background-color:{bg}"]

        display_cols = ["UK Size","US Size","Line","XL Qty","CTN Range","Num CTNs","Qty/CTN","PDF Qty","Diff","Contract OA","Status"]
        existing = [c for c in display_cols if c in size_cmp.columns]
        st.dataframe(
            size_cmp[existing].style.apply(style_size, axis=1),
            use_container_width=True, hide_index=True, height=400
        )

    # Discrepancy detail
    issues = size_cmp[size_cmp["Status"] != "MATCH"]
    if not issues.empty:
        st.warning(f"⚠️ Ditemukan {len(issues)} baris dengan status tidak sempurna:")
        st.dataframe(issues[existing], use_container_width=True, hide_index=True)
    else:
        st.success("✅ Semua size dan quantity cocok antara Excel dan PDF!")

    st.divider()
    st.subheader("⬇️ Download Hasil Perbandingan")

    excel_buf = build_excel_output(xl_hdr, pdf_hdr, hdr_cmp, size_cmp, xl_sizes, pdf_sizes, xl_raw)

    po = xl_hdr.get("PO Number","").replace("/","_")
    art = xl_hdr.get("Article Number","")
    fname = f"comparison_{art}_{po}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

    st.download_button(
        label="📥 Download Excel Report",
        data=excel_buf,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

    with st.expander("🔍 Raw PDF Text (untuk debugging)"):
        st.text(pdf_text[:3000])

elif xl_file and not pdf_file:
    st.info("📄 Silakan upload file PDF Carton Form untuk memulai perbandingan.")
elif pdf_file and not xl_file:
    st.info("📊 Silakan upload file Excel order lines untuk memulai perbandingan.")
else:
    st.info("👆 Upload kedua file di atas untuk memulai perbandingan otomatis.")
    with st.expander("ℹ️ Petunjuk Penggunaan"):
        st.markdown("""
        **Excel file** harus mengandung kolom:
        - `Order #`, `Market PO Number`, `Article Number`, `Model Name`
        - `Manufacturing Size` (UK size), `Customer Size` (US size), `Quantity`
        - `Shipment Method`, `VAS/SHAS L15 – Packing Mode`
        - `Contract Outline Agreement Number`

        **PDF file** adalah Carton Form standar dengan format:
        - Header: PO NO, CUST.O/N, ART. NO, Model, ShipType
        - Tabel size: CTN range, jumlah karton, qty/ctn, total qty, UK size, US size

        **Output Excel** berisi 4 sheet:
        1. **Summary** — KPI & header comparison
        2. **Size Comparison** — detail per size dengan color coding
        3. **Discrepancies** — daftar perbedaan saja
        4. **Raw Excel Data** — data mentah dari file Excel
        """)
