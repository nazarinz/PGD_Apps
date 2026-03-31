import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

st.set_page_config(page_title="Infor vs SAP Carton Comparator", page_icon="📦", layout="wide")

# ── Colors & Style Helpers ────────────────────────────────────────────────────
C = dict(
    GREEN="00C853", RED="D50000", YELLOW="FFD600", BLUE="1565C0",
    ORANGE="E65100", LGRAY="F5F5F5", DGRAY="424242", WHITE="FFFFFF",
    TEAL="00695C", DARKGREEN="1B5E20", LIGHTGREEN="E8F5E9",
    LIGHTRED="FFEBEE", MIDGRAY="9E9E9E",
)

def hfont(color="FFFFFF", sz=10): return Font(name="Arial", bold=True, color=color, size=sz)
def cfont(bold=False, color="000000", sz=10): return Font(name="Arial", bold=bold, color=color, size=sz)
def fill(h): return PatternFill("solid", start_color=h, fgColor=h)
def center(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
def left(): return Alignment(horizontal="left", vertical="center", wrap_text=True)
def bdr():
    s = Side(style="thin", color="BDBDBD")
    return Border(left=s, right=s, top=s, bottom=s)

def wc(ws, r, c, v, bg=None, fnt=None, aln=None, b=True, nf=None):
    cell = ws.cell(row=r, column=c, value=v)
    if bg:  cell.fill = fill(bg)
    if fnt: cell.font = fnt
    if aln: cell.alignment = aln
    if b:   cell.border = bdr()
    if nf:  cell.number_format = nf
    return cell

# ── Excel (Infor) Parsing ────────────────────────────────────────────────────
def load_excel(file):
    df = pd.read_excel(file, dtype=str)
    df.columns = df.columns.str.strip()
    return df

def xl_sizes_for_po(df, po):
    rows = df[df["Order #"].str.strip() == po.strip()].copy()
    out = []
    for _, r in rows.iterrows():
        try: qty = int(float(r.get("Quantity", 0)))
        except: qty = 0
        out.append({
            "UK Size": str(r.get("Manufacturing Size", "")).strip(),
            "US Size": str(r.get("Customer Size", "")).strip(),
            "XL Qty":  qty,
            "Line":    str(r.get("Item Line Number", "")).strip(),
        })
    return pd.DataFrame(out) if out else pd.DataFrame(columns=["UK Size","US Size","XL Qty","Line"])

def xl_hdr_for_po(df, po):
    r = df[df["Order #"].str.strip() == po.strip()]
    if r.empty: return {}
    r = r.iloc[0]
    return {
        "PO Number":   po,
        "Market PO":   str(r.get("Market PO Number", "")),
        "Article":     str(r.get("Article Number", "")),
        "Model":       str(r.get("Model Name", "")),
        "Ship Method": str(r.get("Shipment Method", "")),
        "Pack Mode":   str(r.get("VAS/SHAS L15 – Packing Mode", "")),
        "Destination": str(r.get("FinalDestinationName", "")),
    }

# ── PDF (SAP) Parsing ────────────────────────────────────────────────────────
def _f(pat, text, default=""):
    m = re.search(pat, text)
    return m.group(1).strip() if m else default

def parse_pdf(file, filename=""):
    result = {}
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            po = _f(r"Cust\.PO\s*:\s*(\d+)", text)
            if not po:
                continue
            hdr = {
                "PO Number":   po,
                "Cust Order":  _f(r"CUST\.O/N\s+(\d+)", text),
                "Article":     _f(r"ART\. NO[:\s]+(\w+)", text),
                "Model":       _f(r"Model\s*:\s*([A-Z][A-Z0-9 ]+?)(?:\n|Arr|Ship|Coun|Sub|End|Sale|Cust|OUTER)", text),
                "Ship Method": _f(r"ShipType:\s*\d+\s+(\w+)", text),
                "Pack Mode":   _f(r"(SSP|MSP|USSP|SP)\s+TOTAL", text),
                "Total Pairs": _f(r"Pair[:\s]+([\d,]+)\s+Pairs", text).replace(",", ""),
                "Total CTNs":  _f(r"Ctns[:\s]+([\d,]+)\s+Ctn", text).replace(",", ""),
                "Arr Port":    _f(r"Arr\.Po[:\s]+(\w+)", text),
                "Source File": filename,
            }
            pat = re.compile(
                r"(\d+[-–]\d+|\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+([\d]+[-K]?[-]?)\s+([\d]+[-K]?[-]?)\s+[\d.]+"
            )
            rows = [
                {
                    "CTN Range": m.group(1), "Num CTNs": int(m.group(2)),
                    "Qty/CTN": int(m.group(3)), "PDF Qty": int(m.group(4)),
                    "UK Size": m.group(5).strip(), "US Size": m.group(6).strip(),
                }
                for m in pat.finditer(text)
            ]
            result[po] = {
                "header": hdr,
                "sizes": pd.DataFrame(rows) if rows else pd.DataFrame(),
            }
    return result

# ── Field-by-Field Comparison ─────────────────────────────────────────────────
def compare_po_fields(xl_df, pdf_data, po):
    """
    Returns a DataFrame with columns:
        Field | Infor Value | SAP Value | Status
    Rows: Total Qty (Pairs), then one row per UK Size.
    """
    xs = xl_sizes_for_po(xl_df, po)
    pg = pdf_data.get(po, {})
    ph = pg.get("header", {})
    ps = pg.get("sizes", pd.DataFrame())

    fields = []

    # ── Total Qty ──
    xl_total  = int(xs["XL Qty"].sum()) if not xs.empty else 0
    pdf_total = int(ph.get("Total Pairs", 0) or 0)
    fields.append({
        "Field":       "Total Qty (Pairs)",
        "Infor Value": xl_total,
        "SAP Value":   pdf_total,
        "Status":      "✅ MATCH" if xl_total == pdf_total else "❌ MISMATCH",
    })

    # ── Per-size rows ──
    if not xs.empty and ps is not None and not ps.empty:
        mg = pd.merge(
            xs[["UK Size", "XL Qty"]],
            ps[["UK Size", "PDF Qty"]],
            on="UK Size", how="outer"
        )
    elif not xs.empty:
        mg = xs[["UK Size", "XL Qty"]].copy()
        mg["PDF Qty"] = 0
    elif ps is not None and not ps.empty:
        mg = ps[["UK Size", "PDF Qty"]].copy()
        mg["XL Qty"] = 0
    else:
        mg = pd.DataFrame(columns=["UK Size", "XL Qty", "PDF Qty"])

    if not mg.empty:
        mg["XL Qty"]  = mg["XL Qty"].fillna(0).astype(int)
        mg["PDF Qty"] = mg["PDF Qty"].fillna(0).astype(int)
        # Sort sizes naturally
        mg = mg.sort_values("UK Size").reset_index(drop=True)
        for _, row in mg.iterrows():
            infor_v = int(row["XL Qty"])
            sap_v   = int(row["PDF Qty"])
            fields.append({
                "Field":       f"Qty Size {row['UK Size']}",
                "Infor Value": infor_v,
                "SAP Value":   sap_v,
                "Status":      "✅ MATCH" if infor_v == sap_v else "❌ MISMATCH",
            })

    return pd.DataFrame(fields)

# ── Excel Report ──────────────────────────────────────────────────────────────
def build_report(xl_df, pdf_data, all_pos, xl_filename, pdf_filename):
    wb = Workbook()
    xl_pos_set = set(xl_df["Order #"].str.strip().tolist())

    # ── Sheet 1: PO Compare Summary ──────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "PO Compare Summary"
    ws1.sheet_view.showGridLines = False

    ws1.merge_cells("A1:G1")
    c = ws1["A1"]
    c.value = "PO COMPARE SUMMARY — adidas Infor vs SAP Carton"
    c.font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    c.fill = fill(C["DGRAY"]); c.alignment = center()
    ws1.row_dimensions[1].height = 28

    ws1.merge_cells("A2:G2")
    c = ws1["A2"]
    c.value = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  Infor File: {xl_filename}  |  SAP File: {pdf_filename}"
    c.font = Font(name="Arial", size=9, color="616161")
    c.fill = fill(C["LGRAY"]); c.alignment = left()
    ws1.row_dimensions[2].height = 14

    sum_hdrs = ["PO Number", "Infor File", "SAP File(s)", "Total Fields", "✅ Match", "❌ Mismatch", "Result"]
    for ci, h in enumerate(sum_hdrs, 1):
        wc(ws1, 4, ci, h, bg=C["DGRAY"], fnt=hfont(), aln=center())
    ws1.row_dimensions[4].height = 16

    for ri, po in enumerate(all_pos, 5):
        df_fields = compare_po_fields(xl_df, pdf_data, po)
        total_f   = len(df_fields)
        n_match   = (df_fields["Status"] == "✅ MATCH").sum() if not df_fields.empty else 0
        n_mis     = (df_fields["Status"] == "❌ MISMATCH").sum() if not df_fields.empty else 0

        infor_file = xl_filename
        sap_file   = pdf_data.get(po, {}).get("header", {}).get("Source File", pdf_filename)

        if po not in pdf_data:
            result_txt = "⚠️ NO SAP DATA"
            r_bg, r_fc = C["ORANGE"], "FFFFFF"
        elif po not in xl_pos_set:
            result_txt = "⚠️ NO INFOR DATA"
            r_bg, r_fc = C["ORANGE"], "FFFFFF"
        elif n_mis == 0:
            result_txt = "✅ ALL OK"
            r_bg, r_fc = C["GREEN"], "FFFFFF"
        else:
            result_txt = f"❌ {n_mis} ISSUE(S)"
            r_bg, r_fc = C["RED"], "FFFFFF"

        alt = C["LGRAY"] if ri % 2 == 0 else C["WHITE"]
        ws1.row_dimensions[ri].height = 15
        row_vals = [po, infor_file, sap_file, total_f, n_match, n_mis, result_txt]
        for ci, v in enumerate(row_vals, 1):
            if ci == 7:
                wc(ws1, ri, ci, v, bg=r_bg, fnt=hfont(color=r_fc), aln=center())
            elif ci == 5:
                wc(ws1, ri, ci, v, bg=alt, fnt=cfont(bold=True, color=C["DARKGREEN"], sz=9), aln=center())
            elif ci == 6 and n_mis > 0:
                wc(ws1, ri, ci, v, bg=alt, fnt=cfont(bold=True, color=C["RED"], sz=9), aln=center())
            else:
                wc(ws1, ri, ci, v, bg=alt, fnt=cfont(sz=9), aln=left() if ci <= 3 else center())

    for ci, w in enumerate([16, 36, 36, 13, 10, 13, 18], 1):
        ws1.column_dimensions[get_column_letter(ci)].width = w
    ws1.freeze_panes = "A5"

    # ── Sheet 2: PO Compare Detail ────────────────────────────────────────────
    ws2 = wb.create_sheet("PO Compare Detail")
    ws2.sheet_view.showGridLines = False

    ws2.merge_cells("A1:E1")
    c = ws2["A1"]
    c.value = "PO COMPARE DETAIL — Field by Field"
    c.font = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    c.fill = fill(C["DGRAY"]); c.alignment = center()
    ws2.row_dimensions[1].height = 24

    det_hdrs = ["PO Number", "Field", "Infor Value", "SAP Value", "Status"]
    for ci, h in enumerate(det_hdrs, 1):
        wc(ws2, 2, ci, h, bg="37474F", fnt=hfont(), aln=center())
    ws2.row_dimensions[2].height = 16

    ri2 = 3
    for po in all_pos:
        df_fields = compare_po_fields(xl_df, pdf_data, po)
        if df_fields.empty:
            continue
        for _, row in df_fields.iterrows():
            st = row["Status"]
            is_match = st == "✅ MATCH"
            alt = C["LGRAY"] if ri2 % 2 == 0 else C["WHITE"]
            ws2.row_dimensions[ri2].height = 14
            vals = [po, row["Field"], row["Infor Value"], row["SAP Value"], st]
            for ci, v in enumerate(vals, 1):
                if ci == 5:
                    bg = C["GREEN"] if is_match else C["RED"]
                    fc = "FFFFFF"
                    wc(ws2, ri2, ci, v, bg=bg, fnt=cfont(bold=True, color=fc, sz=9), aln=center())
                elif ci in (3, 4) and not is_match:
                    wc(ws2, ri2, ci, v, bg=C["LIGHTRED"] if C.get("LIGHTRED") else "FFEBEE",
                       fnt=cfont(bold=True, color=C["RED"], sz=9), aln=center())
                elif ci == 1:
                    wc(ws2, ri2, ci, v, bg=alt, fnt=cfont(bold=True, sz=9), aln=left())
                else:
                    wc(ws2, ri2, ci, v, bg=alt, fnt=cfont(sz=9),
                       aln=left() if ci == 2 else center())
            ri2 += 1

    for ci, w in enumerate([16, 22, 14, 14, 14], 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w
    ws2.freeze_panes = "A3"

    # ── Sheet 3: Discrepancies Only ───────────────────────────────────────────
    ws3 = wb.create_sheet("Discrepancies Only")
    ws3.sheet_view.showGridLines = False

    ws3.merge_cells("A1:E1")
    c = ws3["A1"]
    c.value = "DISCREPANCIES ONLY — ❌ MISMATCH rows"
    c.font = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    c.fill = fill(C["ORANGE"]); c.alignment = center()
    ws3.row_dimensions[1].height = 24

    for ci, h in enumerate(det_hdrs, 1):
        wc(ws3, 2, ci, h, bg="37474F", fnt=hfont(), aln=center())
    ws3.row_dimensions[2].height = 16

    ri3 = 3
    for po in all_pos:
        df_fields = compare_po_fields(xl_df, pdf_data, po)
        if df_fields.empty:
            continue
        issues = df_fields[df_fields["Status"] == "❌ MISMATCH"]
        if issues.empty:
            continue
        for _, row in issues.iterrows():
            alt = C["LGRAY"] if ri3 % 2 == 0 else C["WHITE"]
            ws3.row_dimensions[ri3].height = 14
            vals = [po, row["Field"], row["Infor Value"], row["SAP Value"], row["Status"]]
            for ci, v in enumerate(vals, 1):
                if ci == 5:
                    wc(ws3, ri3, ci, v, bg=C["RED"], fnt=cfont(bold=True, color="FFFFFF", sz=9), aln=center())
                elif ci in (3, 4):
                    wc(ws3, ri3, ci, v, bg="FFEBEE", fnt=cfont(bold=True, color=C["RED"], sz=9), aln=center())
                elif ci == 1:
                    wc(ws3, ri3, ci, v, bg=alt, fnt=cfont(bold=True, sz=9), aln=left())
                else:
                    wc(ws3, ri3, ci, v, bg=alt, fnt=cfont(sz=9), aln=left() if ci == 2 else center())
            ri3 += 1

    if ri3 == 3:
        c = ws3.cell(3, 1, "✅ No discrepancies found — all fields match!")
        c.font = Font(name="Arial", bold=True, color=C["GREEN"], size=11)

    for ci, w in enumerate([16, 22, 14, 14, 14], 1):
        ws3.column_dimensions[get_column_letter(ci)].width = w
    ws3.freeze_panes = "A3"

    # ── Sheet 4: Raw Infor Data ───────────────────────────────────────────────
    ws_r = wb.create_sheet("Raw Infor Data")
    ws_r.sheet_view.showGridLines = False
    for ci, col in enumerate(xl_df.columns, 1):
        wc(ws_r, 1, ci, col, bg=C["DGRAY"], fnt=hfont(sz=9), aln=center())
    for ri, row in xl_df.iterrows():
        for ci, v in enumerate(row, 1):
            wc(ws_r, ri + 2, ci, v, fnt=cfont(sz=8), aln=left(),
               bg=C["LGRAY"] if ri % 2 == 0 else C["WHITE"])
    for ci, col in enumerate(xl_df.columns, 1):
        ws_r.column_dimensions[get_column_letter(ci)].width = max(10, min(28, len(str(col)) + 2))
    ws_r.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Streamlit UI ──────────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
    .stApp { font-family: Arial, sans-serif; }
    .block-container { padding-top: 1.2rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📦 Infor vs SAP Carton — Field-by-Field PO Comparator")
st.caption(
    "Match key: **Order # (Infor/Excel)** = **Cust.PO (SAP/PDF)**  |  "
    "Compares Total Qty + per-size quantities field by field, no COA check."
)

col1, col2 = st.columns(2)
with col1:
    st.subheader("📊 Infor File — Order List (Excel)")
    xl_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx", "xls"])
with col2:
    st.subheader("📄 SAP Carton Form (PDF)")
    pdf_file = st.file_uploader("Upload PDF (multi-page)", type=["pdf"])

if xl_file and pdf_file:
    with st.spinner("Processing & matching data..."):
        xl_df    = load_excel(xl_file)
        pdf_data = parse_pdf(pdf_file, filename=pdf_file.name)
        xl_pos   = xl_df["Order #"].str.strip().unique().tolist()
        pdf_pos  = list(pdf_data.keys())
        all_pos  = sorted(set(xl_pos + pdf_pos))
        match_pos = [p for p in pdf_pos if p in xl_pos]
        xl_pos_set = set(xl_pos)

    # ── Build per-PO summary ──
    summary_rows = []
    for po in all_pos:
        df_fields = compare_po_fields(xl_df, pdf_data, po)
        total_f   = len(df_fields)
        n_match   = (df_fields["Status"] == "✅ MATCH").sum() if not df_fields.empty else 0
        n_mis     = (df_fields["Status"] == "❌ MISMATCH").sum() if not df_fields.empty else 0
        sap_file  = pdf_data.get(po, {}).get("header", {}).get("Source File", pdf_file.name)

        if po not in pdf_data:
            result = "⚠️ NO SAP DATA"
        elif po not in xl_pos_set:
            result = "⚠️ NO INFOR DATA"
        elif n_mis == 0:
            result = "✅ ALL OK"
        else:
            result = f"❌ {n_mis} ISSUE(S)"

        summary_rows.append({
            "PO Number":    po,
            "Infor File":   xl_file.name,
            "SAP File(s)":  sap_file,
            "Total Fields": total_f,
            "✅ Match":     n_match,
            "❌ Mismatch":  n_mis,
            "Result":       result,
        })

    summary_df = pd.DataFrame(summary_rows)

    # ── KPI metrics ──
    total_ok   = (summary_df["Result"] == "✅ ALL OK").sum()
    total_mis  = summary_df["❌ Mismatch"].sum()
    total_warn = summary_df["Result"].str.startswith("⚠️").sum()
    pdf_only_list = [p for p in pdf_pos if p not in xl_pos_set]
    xl_only_list  = [p for p in xl_pos  if p not in pdf_pos]

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Matched POs",    len(match_pos))
    k2.metric("✅ All OK",       total_ok)
    k3.metric("❌ POs w/ Issues", (summary_df["❌ Mismatch"] > 0).sum())
    k4.metric("Total Mismatches", int(total_mis))
    k5.metric("SAP only",         len(pdf_only_list))
    k6.metric("Infor only",       len(xl_only_list))

    st.divider()

    # ── Summary Table ──
    st.subheader("📋 PO Compare Summary — adidas Infor vs SAP Carton")

    def _color_result(val):
        if "ALL OK" in str(val):   return "background-color:#c8f7c5; color:#1B5E20; font-weight:bold"
        if "ISSUE" in str(val):    return "background-color:#ffcdd2; color:#B71C1C; font-weight:bold"
        if "NO SAP" in str(val) or "NO INFOR" in str(val): return "background-color:#ffe0b2; color:#E65100; font-weight:bold"
        return ""

    def _color_mismatch(val):
        if isinstance(val, (int, float)) and val > 0:
            return "color:#B71C1C; font-weight:bold"
        return "color:#1B5E20; font-weight:bold"

    styled = (
        summary_df.style
        .map(_color_result, subset=["Result"])
        .map(_color_mismatch, subset=["❌ Mismatch"])
    )
    st.dataframe(styled, use_container_width=True, hide_index=True, height=min(600, 80 + len(summary_df) * 36))

    # ── Detail Table ──
    st.divider()
    st.subheader("🔍 PO Compare Detail — Field by Field")

    detail_rows = []
    for po in all_pos:
        df_fields = compare_po_fields(xl_df, pdf_data, po)
        if df_fields.empty:
            continue
        df_fields.insert(0, "PO Number", po)
        detail_rows.append(df_fields)

    if detail_rows:
        detail_df = pd.concat(detail_rows, ignore_index=True)

        def _color_status(val):
            if val == "✅ MATCH":    return "background-color:#c8f7c5; color:#1B5E20; font-weight:bold"
            if val == "❌ MISMATCH": return "background-color:#ffcdd2; color:#B71C1C; font-weight:bold"
            return ""

        def _color_vals(row):
            if row.get("Status") == "❌ MISMATCH":
                return [""] * 2 + ["background-color:#ffebee; color:#B71C1C; font-weight:bold"] * 2 + [""]
            return [""] * 5

        styled_detail = (
            detail_df.style
            .map(_color_status, subset=["Status"])
            .apply(_color_vals, axis=1, subset=["PO Number", "Field", "Infor Value", "SAP Value", "Status"])
        )
        st.dataframe(styled_detail, use_container_width=True, hide_index=True,
                     height=min(700, 80 + len(detail_df) * 34))
    else:
        st.info("No detail data to display.")

    # ── Discrepancy Summary ──
    st.divider()
    has_issues = summary_df["❌ Mismatch"].sum() > 0
    if has_issues:
        st.subheader("❌ Discrepancies Only")
        disc_rows = []
        for po in all_pos:
            df_fields = compare_po_fields(xl_df, pdf_data, po)
            if df_fields.empty: continue
            iss = df_fields[df_fields["Status"] == "❌ MISMATCH"].copy()
            if iss.empty: continue
            iss.insert(0, "PO Number", po)
            disc_rows.append(iss)
        if disc_rows:
            disc_df = pd.concat(disc_rows, ignore_index=True)

            def _ds(row):
                return [""] * 2 + ["background-color:#ffebee; color:#B71C1C; font-weight:bold"] * 2 + ["background-color:#ffcdd2; color:#B71C1C; font-weight:bold"]

            st.dataframe(
                disc_df.style.apply(_ds, axis=1),
                use_container_width=True, hide_index=True,
                height=min(500, 80 + len(disc_df) * 34)
            )
    else:
        st.success("✅ No discrepancies found — all matched POs are 100% OK!")

    # ── Download ──
    st.divider()
    st.subheader("⬇️ Download Excel Report")
    with st.spinner("Building Excel report..."):
        report_buf = build_report(xl_df, pdf_data, all_pos, xl_file.name, pdf_file.name)
    fname = f"InforVsSAP_Comparison_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button(
        "📥 Download Excel Report",
        data=report_buf,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

    if pdf_only_list:
        st.warning(f"⚠️ {len(pdf_only_list)} PO found in SAP only (not in Infor): {', '.join(pdf_only_list)}")
    if xl_only_list:
        with st.expander(f"ℹ️ {len(xl_only_list)} PO found in Infor only (not in SAP PDF)"):
            st.write(xl_only_list)

elif xl_file and not pdf_file:
    st.info("📄 Please upload the SAP Carton Form PDF.")
elif pdf_file and not xl_file:
    st.info("📊 Please upload the Infor Order List Excel.")
else:
    st.info("👆 Upload both files above to start the comparison.")
    with st.expander("ℹ️ How it works"):
        st.markdown("""
        **Match key**: `Order #` (Infor Excel) = `Cust.PO` (SAP PDF) — matched per PO number

        **Fields compared per PO:**
        - `Total Qty (Pairs)` — sum of all sizes vs PDF total
        - `Qty Size <X>` — per UK size quantity

        **Status per field:**
        - ✅ MATCH — values are identical
        - ❌ MISMATCH — values differ

        **Excel output sheets:**
        1. `PO Compare Summary` — one row per PO with match/mismatch counts
        2. `PO Compare Detail` — all fields for all POs
        3. `Discrepancies Only` — only ❌ MISMATCH rows
        4. `Raw Infor Data` — original Excel upload

        > ℹ️ COA field is **not** compared in this version.
        """)
