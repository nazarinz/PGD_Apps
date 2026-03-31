import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

st.set_page_config(page_title="PO Order vs Carton Form", page_icon="📦", layout="wide")

C = dict(
    GREEN="00C853", RED="D50000", YELLOW="FFD600", BLUE="1565C0",
    ORANGE="E65100", LGRAY="F5F5F5", DGRAY="424242", WHITE="FFFFFF",
    TEAL="00695C"
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

STATUS_STYLE = {
    "MATCH":               (C["GREEN"],  "FFFFFF"),
    "MATCH - COA MISSING": (C["YELLOW"], "000000"),
    "MISMATCH":            (C["RED"],    "FFFFFF"),
    "ONLY IN EXCEL":       (C["ORANGE"], "FFFFFF"),
    "ONLY IN PDF":         (C["ORANGE"], "FFFFFF"),
}

# ── Excel ────────────────────────────────────────────────────────────────────
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
            "UK Size":    str(r.get("Manufacturing Size","")).strip(),
            "US Size":    str(r.get("Customer Size","")).strip(),
            "XL Qty":     qty,
            "Contract OA": str(r.get("Contract Outline Agreement Number","")).strip(),
            "Line":       str(r.get("Item Line Number","")).strip(),
        })
    return pd.DataFrame(out) if out else pd.DataFrame(columns=["UK Size","US Size","XL Qty","Contract OA","Line"])

def xl_hdr_for_po(df, po):
    r = df[df["Order #"].str.strip() == po.strip()]
    if r.empty: return {}
    r = r.iloc[0]
    return {
        "PO Number":  po,
        "Market PO":  str(r.get("Market PO Number","")),
        "Article":    str(r.get("Article Number","")),
        "Model":      str(r.get("Model Name","")),
        "Ship Method":str(r.get("Shipment Method","")),
        "Pack Mode":  str(r.get("VAS/SHAS L15 \u2013 Packing Mode","")),
        "Destination":str(r.get("FinalDestinationName","")),
    }

# ── PDF ──────────────────────────────────────────────────────────────────────
def _f(pat, text, default=""):
    m = re.search(pat, text)
    return m.group(1).strip() if m else default

def parse_pdf(file):
    result = {}
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            po = _f(r"Cust\.PO\s*:\s*(\d+)", text)
            if not po: continue
            hdr = {
                "PO Number":  po,
                "Cust Order": _f(r"CUST\.O/N\s+(\d+)", text),
                "Article":    _f(r"ART\. NO[:\s]+(\w+)", text),
                "Model":      _f(r"Model\s*:\s*([A-Z][A-Z0-9 ]+?)(?:\n|Arr|Ship|Coun|Sub|End|Sale|Cust|OUTER)", text),
                "Ship Method":_f(r"ShipType:\s*\d+\s+(\w+)", text),
                "Pack Mode":  _f(r"(SSP|MSP|USSP|SP)\s+TOTAL", text),
                "Total Pairs":_f(r"Pair[:\s]+([\d,]+)\s+Pairs", text).replace(",",""),
                "Total CTNs": _f(r"Ctns[:\s]+([\d,]+)\s+Ctn", text).replace(",",""),
                "Arr Port":   _f(r"Arr\.Po[:\s]+(\w+)", text),
            }
            pat = re.compile(
                r"(\d+[-\u2013]\d+|\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+([\d]+[-K]?[-]?)\s+([\d]+[-K]?[-]?)\s+[\d.]+"
            )
            rows = [{"CTN Range":m.group(1),"Num CTNs":int(m.group(2)),"Qty/CTN":int(m.group(3)),
                     "PDF Qty":int(m.group(4)),"UK Size":m.group(5).strip(),"US Size":m.group(6).strip()}
                    for m in pat.finditer(text)]
            result[po] = {"header": hdr, "sizes": pd.DataFrame(rows) if rows else pd.DataFrame()}
    return result

# ── Compare ──────────────────────────────────────────────────────────────────
def compare_po(xl_df, pdf_data, po):
    xh  = xl_hdr_for_po(xl_df, po)
    xs  = xl_sizes_for_po(xl_df, po)
    pg  = pdf_data.get(po, {})
    ph  = pg.get("header", {})
    ps  = pg.get("sizes", pd.DataFrame())

    if xs.empty and (ps is None or ps.empty):
        sc = pd.DataFrame()
    elif xs.empty:
        sc = ps.copy(); sc["XL Qty"]=0; sc["Diff"]=-sc["PDF Qty"]
        sc["Status"]="ONLY IN PDF"; sc["Contract OA"]=""; sc["Line"]=""
    elif ps is None or ps.empty:
        sc = xs.copy(); sc["PDF Qty"]=0; sc["Diff"]=sc["XL Qty"]
        sc["Status"]="ONLY IN EXCEL"; sc["CTN Range"]=""; sc["Num CTNs"]=""; sc["Qty/CTN"]=""
    else:
        mg = pd.merge(xs[["UK Size","US Size","XL Qty","Contract OA","Line"]],
                      ps[["UK Size","US Size","PDF Qty","CTN Range","Num CTNs","Qty/CTN"]],
                      on=["UK Size","US Size"], how="outer")
        mg["XL Qty"]  = mg["XL Qty"].fillna(0).astype(int)
        mg["PDF Qty"] = mg["PDF Qty"].fillna(0).astype(int)
        mg["Diff"]    = mg["XL Qty"] - mg["PDF Qty"]
        def _st(row):
            no_c = pd.isna(row.get("CTN Range")) or str(row.get("CTN Range","")).strip() == ""
            no_x = pd.isna(row.get("Line"))      or str(row.get("Line","")).strip() == ""
            if no_c: return "ONLY IN EXCEL"
            if no_x: return "ONLY IN PDF"
            if row["Diff"] != 0: return "MISMATCH"
            coa = str(row.get("Contract OA","")).strip()
            if coa in ("","nan","None"): return "MATCH - COA MISSING"
            return "MATCH"
        mg["Status"] = mg.apply(_st, axis=1)
        sc = mg

    return {"xh": xh, "ph": ph, "sc": sc, "xs": xs, "ps": ps}

# ── Excel Report ──────────────────────────────────────────────────────────────
def build_report(xl_df, pdf_data, all_pos):
    wb = Workbook()

    # Master Summary
    ws = wb.active; ws.title = "Master Summary"
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:L1")
    c = ws["A1"]; c.value = "PO ORDER vs CARTON FORM — MASTER COMPARISON"
    c.font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    c.fill = fill(C["DGRAY"]); c.alignment = center()
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:L2")
    c = ws["A2"]
    c.value = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  PDF POs: {len(pdf_data)}  |  Excel POs: {xl_df['Order #'].nunique()}"
    c.font = Font(name="Arial", size=9, color="616161"); c.fill = fill(C["LGRAY"]); c.alignment = left()
    ws.row_dimensions[2].height = 15

    hdrs = ["PO Number","Article","Model","PDF Pairs","XL Pairs","Diff","PDF CTNs",
            "Sizes OK","Mismatch","COA Warn","Status","Notes"]
    for ci, h in enumerate(hdrs, 1):
        wc(ws, 4, ci, h, bg=C["DGRAY"], fnt=hfont(), aln=center())
    ws.row_dimensions[4].height = 16

    xl_pos_set = set(xl_df["Order #"].str.strip().tolist())
    for ri, po in enumerate(all_pos, 5):
        cmp  = compare_po(xl_df, pdf_data, po)
        sc   = cmp["sc"]; ph = cmp["ph"]; xh = cmp["xh"]
        pdf_p = int(ph.get("Total Pairs","0") or 0)
        xl_p  = int(cmp["xs"]["XL Qty"].sum()) if not cmp["xs"].empty else 0
        diff  = xl_p - pdf_p
        nm = (sc["Status"].str.startswith("MATCH")).sum() if not sc.empty else 0
        ni = (sc["Status"] == "MISMATCH").sum() if not sc.empty else 0
        nw = (sc["Status"] == "MATCH - COA MISSING").sum() if not sc.empty else 0

        if po not in pdf_data:   ov = "NO PDF"
        elif po not in xl_pos_set: ov = "NO EXCEL"
        elif ni > 0: ov = "MISMATCH"
        elif diff != 0: ov = "QTY DIFF"
        elif nw > 0: ov = "COA WARN"
        else: ov = "OK"

        notes = []
        if nw > 0: notes.append(f"COA missing: {nw} line(s)")
        if ni > 0: notes.append(f"Mismatch: {ni} size(s)")
        if diff != 0: notes.append(f"Pairs diff: {diff:+d}")

        art   = ph.get("Article", xh.get("Article",""))
        model = (ph.get("Model","") or xh.get("Model","") or "")[:28]

        ov_colors = {"OK": C["GREEN"], "COA WARN": C["YELLOW"], "MISMATCH": C["RED"],
                     "QTY DIFF": C["RED"], "NO PDF": C["ORANGE"], "NO EXCEL": C["ORANGE"]}
        ov_fc     = {"OK":"FFFFFF","COA WARN":"000000","MISMATCH":"FFFFFF",
                     "QTY DIFF":"FFFFFF","NO PDF":"FFFFFF","NO EXCEL":"FFFFFF"}
        alt = C["LGRAY"] if ri%2==0 else C["WHITE"]
        ws.row_dimensions[ri].height = 15

        row_vals = [po, art, model, pdf_p, xl_p, diff,
                    int(ph.get("Total CTNs",0) or 0), nm, ni, nw, ov, "; ".join(notes)]
        for ci, v in enumerate(row_vals, 1):
            if ci == 11:
                wc(ws, ri, ci, v, bg=ov_colors.get(ov, C["LGRAY"]),
                   fnt=hfont(color=ov_fc.get(ov,"000000")), aln=center())
            elif ci == 6 and diff != 0:
                wc(ws, ri, ci, v, bg=C["RED"], fnt=hfont(), aln=center())
            elif ci == 9 and ni > 0:
                wc(ws, ri, ci, v, bg=C["RED"], fnt=hfont(), aln=center())
            elif ci == 10 and nw > 0:
                wc(ws, ri, ci, v, bg=C["YELLOW"], fnt=cfont(bold=True), aln=center())
            else:
                wc(ws, ri, ci, v, bg=alt, fnt=cfont(sz=9), aln=left() if ci<=3 else center())

    for ci, w in enumerate([16,10,26,12,12,10,10,10,12,10,12,32], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "A5"

    # Size Detail
    ws2 = wb.create_sheet("Size Detail")
    ws2.sheet_view.showGridLines = False
    ws2.merge_cells("A1:N1")
    c = ws2["A1"]; c.value = "SIZE QTY DETAIL — ALL POs"
    c.font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    c.fill = fill(C["DGRAY"]); c.alignment = center()
    ws2.row_dimensions[1].height = 22

    h2 = ["PO Number","Article","Model","UK Size","US Size","XL Line","XL Qty",
          "CTN Range","CTNs","Qty/CTN","PDF Qty","Diff","Contract OA","Status"]
    for ci, h in enumerate(h2, 1):
        wc(ws2, 2, ci, h, bg="37474F", fnt=hfont(), aln=center())
    ws2.row_dimensions[2].height = 16
    ri2 = 3
    for po in all_pos:
        cmp  = compare_po(xl_df, pdf_data, po)
        sc   = cmp["sc"]; ph = cmp["ph"]; xh = cmp["xh"]
        if sc.empty: continue
        art  = ph.get("Article", xh.get("Article",""))
        mdl  = (ph.get("Model","") or xh.get("Model","") or "")[:24]
        for _, row in sc.iterrows():
            st   = row.get("Status","")
            s_bg, s_fc = STATUS_STYLE.get(st, (C["WHITE"],"000000"))
            alt  = C["LGRAY"] if ri2%2==0 else C["WHITE"]
            ws2.row_dimensions[ri2].height = 14
            vals = [po, art, mdl,
                    row.get("UK Size",""), row.get("US Size",""), row.get("Line",""),
                    int(row.get("XL Qty",0)), row.get("CTN Range",""),
                    row.get("Num CTNs",""), row.get("Qty/CTN",""),
                    int(row.get("PDF Qty",0)), int(row.get("Diff",0)),
                    row.get("Contract OA",""), st]
            for ci, v in enumerate(vals, 1):
                if ci == 14:
                    wc(ws2, ri2, ci, v, bg=s_bg, fnt=cfont(bold=True, color=s_fc, sz=9), aln=center())
                else:
                    wc(ws2, ri2, ci, v, bg=alt, fnt=cfont(sz=9), aln=left() if ci<=3 else center())
            ri2 += 1

    for ci, w in enumerate([14,10,22,9,9,9,10,13,8,8,10,8,18,20], 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w
    ws2.freeze_panes = "A3"

    # Discrepancies
    ws3 = wb.create_sheet("Discrepancies")
    ws3.sheet_view.showGridLines = False
    ws3.merge_cells("A1:N1")
    c = ws3["A1"]; c.value = "DISCREPANCIES ONLY"
    c.font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    c.fill = fill(C["ORANGE"]); c.alignment = center()
    ws3.row_dimensions[1].height = 22
    for ci, h in enumerate(h2, 1):
        wc(ws3, 2, ci, h, bg="37474F", fnt=hfont(), aln=center())
    ws3.row_dimensions[2].height = 16
    ri3 = 3
    for po in all_pos:
        cmp  = compare_po(xl_df, pdf_data, po)
        sc   = cmp["sc"]; ph = cmp["ph"]; xh = cmp["xh"]
        if sc.empty: continue
        issues = sc[sc["Status"] != "MATCH"]
        if issues.empty: continue
        art = ph.get("Article", xh.get("Article",""))
        mdl = (ph.get("Model","") or xh.get("Model","") or "")[:24]
        for _, row in issues.iterrows():
            st   = row.get("Status","")
            s_bg, s_fc = STATUS_STYLE.get(st, (C["YELLOW"],"000000"))
            alt  = C["LGRAY"] if ri3%2==0 else C["WHITE"]
            ws3.row_dimensions[ri3].height = 14
            vals = [po, art, mdl,
                    row.get("UK Size",""), row.get("US Size",""), row.get("Line",""),
                    int(row.get("XL Qty",0)), row.get("CTN Range",""),
                    row.get("Num CTNs",""), row.get("Qty/CTN",""),
                    int(row.get("PDF Qty",0)), int(row.get("Diff",0)),
                    row.get("Contract OA",""), st]
            for ci, v in enumerate(vals, 1):
                if ci == 14:
                    wc(ws3, ri3, ci, v, bg=s_bg, fnt=cfont(bold=True, color=s_fc, sz=9), aln=center())
                else:
                    wc(ws3, ri3, ci, v, bg=alt, fnt=cfont(sz=9), aln=left() if ci<=3 else center())
            ri3 += 1

    if ri3 == 3:
        ws3["A3"] = "No discrepancies found."
        ws3["A3"].font = Font(name="Arial", bold=True, color=C["GREEN"], size=11)

    for ci, w in enumerate([14,10,22,9,9,9,10,13,8,8,10,8,18,20], 1):
        ws3.column_dimensions[get_column_letter(ci)].width = w
    ws3.freeze_panes = "A3"

    # Per-PO sheets
    for po in sorted(pdf_data.keys())[:30]:
        if po not in xl_pos_set: continue
        cmp  = compare_po(xl_df, pdf_data, po)
        sc   = cmp["sc"]; ph = cmp["ph"]; xh = cmp["xh"]
        art  = ph.get("Article", xh.get("Article",""))
        sname = f"PO_{po[-7:]}"
        wsp = wb.create_sheet(sname)
        wsp.sheet_view.showGridLines = False

        wsp.merge_cells("A1:K1")
        c = wsp["A1"]
        c.value = f"PO: {po}  |  {art}  |  {(ph.get('Model','') or xh.get('Model',''))[:30]}"
        c.font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
        c.fill = fill(C["TEAL"]); c.alignment = left()
        wsp.row_dimensions[1].height = 20

        info_items = [
            ("Cust.PO", ph.get("PO Number","")), ("Cust O/N", ph.get("Cust Order","")),
            ("Article", art), ("Ship Method", ph.get("Ship Method", xh.get("Ship Method",""))),
            ("Pack Mode", ph.get("Pack Mode", xh.get("Pack Mode",""))),
            ("PDF Pairs", ph.get("Total Pairs","")), ("PDF CTNs", ph.get("Total CTNs","")),
            ("XL Pairs", str(int(cmp["xs"]["XL Qty"].sum())) if not cmp["xs"].empty else "0"),
            ("Arr Port", ph.get("Arr Port","")),
        ]
        for ii, (lbl, val) in enumerate(info_items):
            rn = 2 + (ii // 5); cn = (ii % 5)*2 + 1
            wc(wsp, rn, cn, lbl, bg=C["DGRAY"], fnt=hfont(sz=9), aln=center())
            wc(wsp, rn, cn+1, val, bg=C["LGRAY"], fnt=cfont(sz=9), aln=left())
        for r in [2,3]: wsp.row_dimensions[r].height = 14

        sz_h = ["UK Size","US Size","XL Line","XL Qty","CTN Range","CTNs","Qty/CTN","PDF Qty","Diff","Contract OA","Status"]
        for ci, h in enumerate(sz_h, 1):
            wc(wsp, 5, ci, h, bg="37474F", fnt=hfont(sz=9), aln=center())
        wsp.row_dimensions[5].height = 14

        if sc.empty:
            wsp.cell(6,1).value = "No data."
        else:
            for rr, (_, row) in enumerate(sc.iterrows(), 6):
                st   = row.get("Status","")
                s_bg, s_fc = STATUS_STYLE.get(st, (C["WHITE"],"000000"))
                alt  = C["LGRAY"] if rr%2==0 else C["WHITE"]
                wsp.row_dimensions[rr].height = 13
                vals = [row.get("UK Size",""), row.get("US Size",""), row.get("Line",""),
                        int(row.get("XL Qty",0)), row.get("CTN Range",""),
                        row.get("Num CTNs",""), row.get("Qty/CTN",""),
                        int(row.get("PDF Qty",0)), int(row.get("Diff",0)),
                        row.get("Contract OA",""), st]
                for ci, v in enumerate(vals, 1):
                    if ci == 11:
                        wc(wsp, rr, ci, v, bg=s_bg, fnt=cfont(bold=True, color=s_fc, sz=9), aln=center())
                    elif ci == 9 and int(v) != 0:
                        wc(wsp, rr, ci, v, bg=C["RED"], fnt=hfont(sz=9), aln=center())
                    else:
                        wc(wsp, rr, ci, v, bg=alt, fnt=cfont(sz=9), aln=center() if ci>2 else left())

        for ci, w in enumerate([9,9,9,11,13,8,8,11,9,20,20], 1):
            wsp.column_dimensions[get_column_letter(ci)].width = w

    # Raw Excel
    ws_r = wb.create_sheet("Raw Excel")
    ws_r.sheet_view.showGridLines = False
    for ci, col in enumerate(xl_df.columns, 1):
        wc(ws_r, 1, ci, col, bg=C["DGRAY"], fnt=hfont(sz=9), aln=center())
    for ri, row in xl_df.iterrows():
        for ci, v in enumerate(row, 1):
            wc(ws_r, ri+2, ci, v, fnt=cfont(sz=8), aln=left(),
               bg=C["LGRAY"] if ri%2==0 else C["WHITE"])
    for ci, col in enumerate(xl_df.columns, 1):
        ws_r.column_dimensions[get_column_letter(ci)].width = max(10, min(28, len(str(col))+2))
    ws_r.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── UI ────────────────────────────────────────────────────────────────────────
st.markdown("<style>.stApp{font-family:Arial,sans-serif}.block-container{padding-top:1.2rem}</style>",
            unsafe_allow_html=True)

st.title("📦 PO Order vs Carton Form — Multi-PO Comparator")
st.caption("Match key: **Order # (Excel)** = **Cust.PO (PDF)**. PDF bisa berisi banyak halaman/PO.")

col1, col2 = st.columns(2)
with col1:
    st.subheader("📊 Excel — Order List")
    xl_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx","xls"])
with col2:
    st.subheader("📄 PDF — Carton Form(s)")
    pdf_file = st.file_uploader("Upload PDF (multi-halaman)", type=["pdf"])

if xl_file and pdf_file:
    with st.spinner("Memproses & mencocokkan data..."):
        xl_df    = load_excel(xl_file)
        pdf_data = parse_pdf(pdf_file)
        xl_pos   = xl_df["Order #"].str.strip().unique().tolist()
        pdf_pos  = list(pdf_data.keys())
        all_pos  = sorted(set(xl_pos + pdf_pos))
        match_pos = [p for p in pdf_pos if p in xl_pos]

    # KPI
    status_list = []
    for po in match_pos:
        cmp = compare_po(xl_df, pdf_data, po)
        sc  = cmp["sc"]
        if sc.empty: status_list.append("OK")
        elif (sc["Status"] == "MISMATCH").any(): status_list.append("MISMATCH")
        elif (sc["Status"] == "MATCH - COA MISSING").any(): status_list.append("WARN")
        else: status_list.append("OK")

    po_ok   = status_list.count("OK")
    po_mis  = status_list.count("MISMATCH")
    po_warn = status_list.count("WARN")
    pdf_only_list = [p for p in pdf_pos if p not in xl_pos]
    xl_only_list  = [p for p in xl_pos  if p not in pdf_pos]

    k1,k2,k3,k4,k5,k6 = st.columns(6)
    k1.metric("Matched POs",    len(match_pos))
    k2.metric("✅ All OK",       po_ok)
    k3.metric("❌ Mismatch",     po_mis)
    k4.metric("⚠️ COA Warning",  po_warn)
    k5.metric("PDF only",        len(pdf_only_list))
    k6.metric("Excel only",      len(xl_only_list))

    st.divider()
    st.subheader("📋 Summary per PO")

    rows = []
    for po in all_pos:
        cmp  = compare_po(xl_df, pdf_data, po)
        sc   = cmp["sc"]; ph = cmp["ph"]; xh = cmp["xh"]
        pdf_p = int(ph.get("Total Pairs","0") or 0)
        xl_p  = int(cmp["xs"]["XL Qty"].sum()) if not cmp["xs"].empty else 0
        diff  = xl_p - pdf_p
        ni = (sc["Status"] == "MISMATCH").sum() if not sc.empty else 0
        nw = (sc["Status"] == "MATCH - COA MISSING").sum() if not sc.empty else 0
        xl_pos_set = set(xl_pos)
        if po not in pdf_data:   ov = "NO PDF"
        elif po not in xl_pos_set: ov = "NO EXCEL"
        elif ni > 0: ov = "MISMATCH"
        elif diff != 0: ov = "QTY DIFF"
        elif nw > 0: ov = "COA WARN"
        else: ov = "OK"
        rows.append({"PO":po,"Article":ph.get("Article",xh.get("Article",""))[:12],
                     "Model":(ph.get("Model","") or xh.get("Model","") or "")[:24],
                     "PDF Pairs":pdf_p,"XL Pairs":xl_p,"Diff":diff,
                     "Mismatch":ni,"COA Warn":nw,"Status":ov})

    summary_df = pd.DataFrame(rows)
    def _cs(val):
        cm = {"OK":"background-color:#c8f7c5","MISMATCH":"background-color:#ffcdd2",
              "COA WARN":"background-color:#fff9c4","QTY DIFF":"background-color:#ffcdd2",
              "NO PDF":"background-color:#ffe0cc","NO EXCEL":"background-color:#ffe0cc"}
        return cm.get(val,"")
    st.dataframe(summary_df.style.map(_cs, subset=["Status"]),
                 use_container_width=True, hide_index=True, height=500)

    # Discrepancy detail
    st.divider()
    st.subheader("❌ Discrepancy Detail")
    disc = []
    for po in match_pos:
        cmp = compare_po(xl_df, pdf_data, po)
        sc  = cmp["sc"]; ph = cmp["ph"]; xh = cmp["xh"]
        if sc.empty: continue
        iss = sc[sc["Status"] != "MATCH"].copy()
        if iss.empty: continue
        iss.insert(0,"PO",po)
        iss.insert(1,"Article",ph.get("Article",xh.get("Article","")))
        disc.append(iss)
    if disc:
        disc_df = pd.concat(disc, ignore_index=True)
        def _cd(row):
            cm = {"MISMATCH":"background-color:#ffcdd2",
                  "MATCH - COA MISSING":"background-color:#fff9c4",
                  "ONLY IN EXCEL":"background-color:#ffe0cc",
                  "ONLY IN PDF":"background-color:#ffe0cc"}
            return [cm.get(row.get("Status",""),"")] * len(row)
        st.dataframe(disc_df.style.apply(_cd, axis=1),
                     use_container_width=True, hide_index=True, height=400)
    else:
        st.success("✅ Tidak ada discrepancy pada PO yang matched!")

    # Download
    st.divider()
    st.subheader("⬇️ Download Report")
    with st.spinner("Membuat Excel report..."):
        report_buf = build_report(xl_df, pdf_data, all_pos)
    fname = f"PO_Comparison_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button("📥 Download Excel Report", data=report_buf, file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True, type="primary")

    if pdf_only_list:
        st.warning(f"⚠️ {len(pdf_only_list)} PO ada di PDF tapi tidak di Excel: {', '.join(pdf_only_list)}")
    if xl_only_list:
        with st.expander(f"ℹ️ {len(xl_only_list)} PO hanya di Excel (tidak ada di PDF)"):
            st.write(xl_only_list)

elif xl_file and not pdf_file:
    st.info("📄 Silakan upload PDF Carton Form.")
elif pdf_file and not xl_file:
    st.info("📊 Silakan upload Excel Order List.")
else:
    st.info("👆 Upload kedua file di atas untuk memulai.")
    with st.expander("ℹ️ Cara Kerja"):
        st.markdown("""
        - **Match key**: `Order #` (Excel) = `Cust.PO` (PDF) — per PO
        - PDF bisa multi-halaman, tiap halaman = 1 PO
        - **Output Excel**: Master Summary + Size Detail + Discrepancies + sheet per-PO + Raw Excel
        - 🟢 OK | 🟡 COA WARN | 🔴 MISMATCH/QTY DIFF | 🟠 NO PDF/NO EXCEL
        """)
