import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

st.set_page_config(page_title="Infor vs SAP Comparator", page_icon="\U0001f4e6", layout="wide")

# ── Colors & Style Helpers ────────────────────────────────────────────────────
C = dict(
    GREEN="00C853", RED="D50000", ORANGE="E65100",
    LGRAY="F5F5F5", DGRAY="424242", WHITE="FFFFFF",
    TEAL="00695C", DARKGREEN="1B5E20", BLUE_HDR="1A237E",
    INFO_BG="E3F2FD", INFO_FG="0D47A1",
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

# ── Excel (Infor) Parsing ─────────────────────────────────────────────────────
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
            "Infor Qty": qty,
            "XL Line":   str(r.get("Item Line Number", "")).strip(),
        })
    return pd.DataFrame(out) if out else pd.DataFrame(
        columns=["UK Size", "US Size", "Infor Qty", "XL Line"])

def xl_hdr_for_po(df, po):
    r = df[df["Order #"].str.strip() == po.strip()]
    if r.empty: return {}
    r = r.iloc[0]
    return {
        "PO Number":   po,
        "Article":     str(r.get("Article Number", "")).strip(),
        "Model":       str(r.get("Model Name", "")).strip(),
        "Ship Method": str(r.get("Shipment Method", "")),
        "Pack Mode":   str(r.get("VAS/SHAS L15 \u2013 Packing Mode", "")),
        "Destination": str(r.get("FinalDestinationName", "")),
    }

# ── PDF (SAP) Parsing ─────────────────────────────────────────────────────────
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
                "Article":     _f(r"ART\. NO[:\s]+(\w+)", text).strip(),
                "Model":       _f(r"Model\s*:\s*([A-Z][A-Z0-9 ]+?)(?:\n|Arr|Ship|Coun|Sub|End|Sale|Cust|OUTER)", text).strip(),
                "Ship Method": _f(r"ShipType:\s*\d+\s+(\w+)", text),
                "Pack Mode":   _f(r"(SSP|MSP|USSP|SP)\s+TOTAL", text),
                "Total Pairs": _f(r"Pair[:\s]+([\d,]+)\s+Pairs", text).replace(",", ""),
                "Total CTNs":  _f(r"Ctns[:\s]+([\d,]+)\s+Ctn", text).replace(",", ""),
                "Arr Port":    _f(r"Arr\.Po[:\s]+(\w+)", text),
                "Source File": filename,
            }
            pat = re.compile(
                r"(\d+[-\u2013]\d+|\d+)\s+(\d+)\s+(\d+)\s+(\d+)"
                r"\s+([\d]+[-K]?[-]?)\s+([\d]+[-K]?[-]?)\s+[\d.]+"
            )
            rows = [
                {
                    "CTN Range": m.group(1),
                    "CTNs":      int(m.group(2)),
                    "Qty/CTN":   int(m.group(3)),
                    "SAP Qty":   int(m.group(4)),
                    "UK Size":   m.group(5).strip(),
                    "US Size":   m.group(6).strip(),
                }
                for m in pat.finditer(text)
            ]
            result[po] = {
                "header": hdr,
                "sizes":  pd.DataFrame(rows) if rows else pd.DataFrame(),
            }
    return result

# ── Field-by-Field Comparison (Qty only) ─────────────────────────────────────
def compare_po_fields(xl_df, pdf_data, po):
    """
    Returns DataFrame: Field | Infor Value | SAP Value | Status
    Fields: Total Qty (Pairs) + Qty per UK Size.
    Article & Model are NOT compared (info-only — handled separately).
    """
    xs = xl_sizes_for_po(xl_df, po)
    pg = pdf_data.get(po, {})
    ph = pg.get("header", {})
    ps = pg.get("sizes", pd.DataFrame())

    rows = []

    # ── Total Qty (Pairs) ──
    infor_total = int(xs["Infor Qty"].sum()) if not xs.empty else 0
    sap_total   = int(ph.get("Total Pairs", 0) or 0)
    rows.append({
        "Field":       "Total Qty (Pairs)",
        "Infor Value": infor_total,
        "SAP Value":   sap_total,
        "Status":      "✅ MATCH" if infor_total == sap_total else "❌ MISMATCH",
    })

    # ── Per-size Qty rows ──
    if not xs.empty and ps is not None and not ps.empty:
        mg = pd.merge(xs[["UK Size", "Infor Qty"]], ps[["UK Size", "SAP Qty"]],
                      on="UK Size", how="outer")
    elif not xs.empty:
        mg = xs[["UK Size", "Infor Qty"]].copy(); mg["SAP Qty"] = 0
    elif ps is not None and not ps.empty:
        mg = ps[["UK Size", "SAP Qty"]].copy(); mg["Infor Qty"] = 0
    else:
        mg = pd.DataFrame(columns=["UK Size", "Infor Qty", "SAP Qty"])

    if not mg.empty:
        mg["Infor Qty"] = mg["Infor Qty"].fillna(0).astype(int)
        mg["SAP Qty"]   = mg["SAP Qty"].fillna(0).astype(int)
        mg = mg.sort_values("UK Size").reset_index(drop=True)
        for _, row in mg.iterrows():
            iv, sv = int(row["Infor Qty"]), int(row["SAP Qty"])
            rows.append({
                "Field":       f"Qty Size {row['UK Size']}",
                "Infor Value": iv,
                "SAP Value":   sv,
                "Status":      "✅ MATCH" if iv == sv else "❌ MISMATCH",
            })

    return pd.DataFrame(rows)

# ── Size Detail (all columns) ─────────────────────────────────────────────────
def compare_po_size_detail(xl_df, pdf_data, po):
    """
    Returns one row per UK Size with columns:
    UK Size | US Size | XL Line | Infor Qty | CTN Range | CTNs | Qty/CTN | SAP Qty | Diff | Status
    Article & Model are added by caller.
    """
    xs = xl_sizes_for_po(xl_df, po)
    pg = pdf_data.get(po, {})
    ps = pg.get("sizes", pd.DataFrame())

    COLS = ["UK Size", "US Size", "XL Line", "Infor Qty",
            "CTN Range", "CTNs", "Qty/CTN", "SAP Qty", "Diff", "Status"]

    if not xs.empty and ps is not None and not ps.empty:
        xl_c  = xs[["UK Size", "US Size", "Infor Qty", "XL Line"]]
        sap_c = ps[["UK Size", "US Size", "SAP Qty", "CTN Range", "CTNs", "Qty/CTN"]].rename(
            columns={"US Size": "_sap_us"})
        mg = pd.merge(xl_c, sap_c, on="UK Size", how="outer")
        # Fill US Size: prefer Infor, fall back to SAP
        mg["US Size"] = mg["US Size"].combine_first(mg["_sap_us"])
        mg = mg.drop(columns=["_sap_us"], errors="ignore")
    elif not xs.empty:
        mg = xs[["UK Size", "US Size", "Infor Qty", "XL Line"]].copy()
        mg["SAP Qty"] = 0
        for col in ["CTN Range", "CTNs", "Qty/CTN"]:
            mg[col] = ""
    elif ps is not None and not ps.empty:
        mg = ps[["UK Size", "US Size", "SAP Qty", "CTN Range", "CTNs", "Qty/CTN"]].copy()
        mg["Infor Qty"] = 0; mg["XL Line"] = ""
    else:
        return pd.DataFrame(columns=COLS)

    mg["Infor Qty"] = mg["Infor Qty"].fillna(0).astype(int)
    mg["SAP Qty"]   = mg["SAP Qty"].fillna(0).astype(int)
    mg["Diff"]      = mg["Infor Qty"] - mg["SAP Qty"]

    def _st(row):
        no_sap   = str(row.get("CTN Range", "")).strip() in ("", "nan", "None", "NaN")
        no_infor = str(row.get("XL Line",   "")).strip() in ("", "nan", "None", "NaN")
        if no_sap:   return "ONLY IN INFOR"
        if no_infor: return "ONLY IN SAP"
        return "✅ MATCH" if row["Diff"] == 0 else "❌ MISMATCH"

    mg["Status"] = mg.apply(_st, axis=1)
    for col in COLS:
        if col not in mg.columns:
            mg[col] = ""
    return mg[COLS].sort_values("UK Size").reset_index(drop=True)

# ── PO Info (Article & Model, no comparison) ─────────────────────────────────
def po_info(xl_df, pdf_data, po):
    xh = xl_hdr_for_po(xl_df, po)
    ph = pdf_data.get(po, {}).get("header", {})
    return {
        "Article (Infor)": xh.get("Article", ""),
        "Article (SAP)":   ph.get("Article", ""),
        "Model (Infor)":   xh.get("Model",   ""),
        "Model (SAP)":     ph.get("Model",   ""),
    }

# ── Excel Report ──────────────────────────────────────────────────────────────
def build_report(xl_df, pdf_data, all_pos, xl_filename, pdf_filename):
    wb = Workbook()
    xl_pos_set = set(xl_df["Order #"].str.strip().tolist())

    ST_STYLE = {
        "✅ MATCH":      (C["GREEN"],  "FFFFFF"),
        "❌ MISMATCH":   (C["RED"],    "FFFFFF"),
        "ONLY IN INFOR": (C["ORANGE"], "FFFFFF"),
        "ONLY IN SAP":   (C["ORANGE"], "FFFFFF"),
    }

    # ─────────────────────────────────────────────────────────────────────────
    # Sheet 1 — PO Compare Summary
    # ─────────────────────────────────────────────────────────────────────────
    ws1 = wb.active; ws1.title = "PO Compare Summary"
    ws1.sheet_view.showGridLines = False

    ws1.merge_cells("A1:G1")
    c = ws1["A1"]
    c.value = "PO COMPARE SUMMARY \u2014 adidas Infor vs SAP Carton"
    c.font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    c.fill = fill(C["DGRAY"]); c.alignment = center(); ws1.row_dimensions[1].height = 28

    ws1.merge_cells("A2:G2")
    c = ws1["A2"]
    c.value = (
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  "
        f"Infor: {xl_filename}  |  SAP: {pdf_filename}"
    )
    c.font = Font(name="Arial", size=9, color="616161")
    c.fill = fill(C["LGRAY"]); c.alignment = left(); ws1.row_dimensions[2].height = 14

    # note row
    ws1.merge_cells("A3:G3")
    c = ws1["A3"]
    c.value = "\u2139\ufe0f  Article & Model shown as info only. Comparison: Total Qty (Pairs) + Qty per Size."
    c.font = Font(name="Arial", italic=True, size=9, color=C["INFO_FG"])
    c.fill = fill(C["INFO_BG"]); c.alignment = left(); ws1.row_dimensions[3].height = 14

    hdrs = ["PO Number", "Infor File", "SAP File(s)", "Total Fields", "✅ Match", "❌ Mismatch", "Result"]
    for ci, h in enumerate(hdrs, 1):
        wc(ws1, 5, ci, h, bg=C["DGRAY"], fnt=hfont(), aln=center())
    ws1.row_dimensions[5].height = 16

    for ri, po in enumerate(all_pos, 6):
        df_f    = compare_po_fields(xl_df, pdf_data, po)
        total_f = len(df_f)
        n_m     = (df_f["Status"] == "✅ MATCH").sum()    if not df_f.empty else 0
        n_mm    = (df_f["Status"] == "❌ MISMATCH").sum() if not df_f.empty else 0
        sap_src = pdf_data.get(po, {}).get("header", {}).get("Source File", pdf_filename)

        if po not in pdf_data:       r_txt, r_bg, r_fc = "\u26a0\ufe0f NO SAP DATA",   C["ORANGE"], "FFFFFF"
        elif po not in xl_pos_set:   r_txt, r_bg, r_fc = "\u26a0\ufe0f NO INFOR DATA", C["ORANGE"], "FFFFFF"
        elif n_mm == 0:              r_txt, r_bg, r_fc = "\u2705 ALL OK",               C["GREEN"],  "FFFFFF"
        else:                        r_txt, r_bg, r_fc = f"\u274c {n_mm} ISSUE(S)",     C["RED"],    "FFFFFF"

        alt = C["LGRAY"] if ri % 2 == 0 else C["WHITE"]
        ws1.row_dimensions[ri].height = 15
        for ci, v in enumerate([po, xl_filename, sap_src, total_f, n_m, n_mm, r_txt], 1):
            if ci == 7:
                wc(ws1, ri, ci, v, bg=r_bg, fnt=hfont(color=r_fc), aln=center())
            elif ci == 5:
                wc(ws1, ri, ci, v, bg=alt, fnt=cfont(bold=True, color=C["DARKGREEN"], sz=9), aln=center())
            elif ci == 6 and n_mm > 0:
                wc(ws1, ri, ci, v, bg=alt, fnt=cfont(bold=True, color=C["RED"], sz=9), aln=center())
            else:
                wc(ws1, ri, ci, v, bg=alt, fnt=cfont(sz=9), aln=left() if ci <= 3 else center())

    for ci, w in enumerate([16, 36, 36, 13, 10, 13, 18], 1):
        ws1.column_dimensions[get_column_letter(ci)].width = w
    ws1.freeze_panes = "A6"

    # ─────────────────────────────────────────────────────────────────────────
    # Sheet 2 — PO Compare Detail (field-by-field, qty only)
    # ─────────────────────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("PO Compare Detail")
    ws2.sheet_view.showGridLines = False

    ws2.merge_cells("A1:E1")
    c = ws2["A1"]
    c.value = "PO COMPARE DETAIL \u2014 Field by Field"
    c.font = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    c.fill = fill(C["DGRAY"]); c.alignment = center(); ws2.row_dimensions[1].height = 24

    ws2.merge_cells("A2:E2")
    c = ws2["A2"]
    c.value = "\u2139\ufe0f  Fields compared: Total Qty (Pairs) + Qty per Size. Article & Model are info-only (see Size Detail sheet)."
    c.font = Font(name="Arial", italic=True, size=9, color=C["INFO_FG"])
    c.fill = fill(C["INFO_BG"]); c.alignment = left(); ws2.row_dimensions[2].height = 14

    det_hdrs = ["PO Number", "Field", "Infor Value", "SAP Value", "Status"]
    for ci, h in enumerate(det_hdrs, 1):
        wc(ws2, 3, ci, h, bg="37474F", fnt=hfont(), aln=center())
    ws2.row_dimensions[3].height = 16

    ri2 = 4
    for po in all_pos:
        df_f = compare_po_fields(xl_df, pdf_data, po)
        if df_f.empty: continue
        for _, row in df_f.iterrows():
            st = row["Status"]; is_m = (st == "✅ MATCH")
            alt = C["LGRAY"] if ri2 % 2 == 0 else C["WHITE"]
            ws2.row_dimensions[ri2].height = 14
            for ci, v in enumerate([po, row["Field"], row["Infor Value"], row["SAP Value"], st], 1):
                if ci == 5:
                    wc(ws2, ri2, ci, v,
                       bg=C["GREEN"] if is_m else C["RED"],
                       fnt=cfont(bold=True, color="FFFFFF", sz=9), aln=center())
                elif ci in (3, 4) and not is_m:
                    wc(ws2, ri2, ci, v, bg="FFEBEE",
                       fnt=cfont(bold=True, color=C["RED"], sz=9), aln=center())
                elif ci == 1:
                    wc(ws2, ri2, ci, v, bg=alt, fnt=cfont(bold=True, sz=9), aln=left())
                else:
                    wc(ws2, ri2, ci, v, bg=alt, fnt=cfont(sz=9),
                       aln=left() if ci == 2 else center())
            ri2 += 1

    for ci, w in enumerate([16, 24, 16, 16, 14], 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w
    ws2.freeze_panes = "A4"

    # ─────────────────────────────────────────────────────────────────────────
    # Sheet 3 — Size Detail (all columns incl. Article, Model as info)
    # ─────────────────────────────────────────────────────────────────────────
    ws3 = wb.create_sheet("Size Detail")
    ws3.sheet_view.showGridLines = False

    ws3.merge_cells("A1:M1")
    c = ws3["A1"]
    c.value = "SIZE DETAIL \u2014 All POs (Article & Model = info only, not compared)"
    c.font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    c.fill = fill(C["DGRAY"]); c.alignment = left(); ws3.row_dimensions[1].height = 20

    sd_hdrs = [
        "PO Number", "Article (Infor)", "Article (SAP)", "Model (Infor)", "Model (SAP)",
        "UK Size", "US Size", "XL Line",
        "Infor Qty", "CTN Range", "CTNs", "Qty/CTN", "SAP Qty", "Diff", "Status",
    ]
    for ci, h in enumerate(sd_hdrs, 1):
        wc(ws3, 2, ci, h, bg="37474F", fnt=hfont(), aln=center())
    ws3.row_dimensions[2].height = 16

    ri3 = 3
    for po in all_pos:
        info = po_info(xl_df, pdf_data, po)
        sd   = compare_po_size_detail(xl_df, pdf_data, po)
        if sd.empty: continue
        for _, row in sd.iterrows():
            st      = row.get("Status", "")
            s_bg, s_fc = ST_STYLE.get(st, (C["WHITE"], "000000"))
            alt     = C["LGRAY"] if ri3 % 2 == 0 else C["WHITE"]
            ws3.row_dimensions[ri3].height = 14
            try:    diff_i = int(row.get("Diff", 0))
            except: diff_i = 0
            vals = [
                po,
                info["Article (Infor)"], info["Article (SAP)"],
                info["Model (Infor)"],   info["Model (SAP)"],
                row.get("UK Size",""),   row.get("US Size",""),
                row.get("XL Line",""),
                int(row.get("Infor Qty", 0)),
                row.get("CTN Range",""), row.get("CTNs",""), row.get("Qty/CTN",""),
                int(row.get("SAP Qty", 0)), diff_i, st,
            ]
            for ci, v in enumerate(vals, 1):
                if ci == 15:
                    wc(ws3, ri3, ci, v, bg=s_bg, fnt=cfont(bold=True, color=s_fc, sz=9), aln=center())
                elif ci == 14 and diff_i != 0:
                    wc(ws3, ri3, ci, v, bg=C["RED"], fnt=hfont(sz=9), aln=center())
                # Article & Model columns — info style (light blue)
                elif ci in (2, 3, 4, 5):
                    wc(ws3, ri3, ci, v, bg=C["INFO_BG"],
                       fnt=cfont(sz=9, color=C["INFO_FG"]), aln=left())
                else:
                    wc(ws3, ri3, ci, v, bg=alt, fnt=cfont(sz=9),
                       aln=left() if ci <= 5 else center())
            ri3 += 1

    for ci, w in enumerate([14, 12, 12, 22, 22, 9, 9, 9, 10, 13, 8, 9, 10, 8, 16], 1):
        ws3.column_dimensions[get_column_letter(ci)].width = w
    ws3.freeze_panes = "A3"

    # ─────────────────────────────────────────────────────────────────────────
    # Sheet 4 — Discrepancies Only
    # ─────────────────────────────────────────────────────────────────────────
    ws4 = wb.create_sheet("Discrepancies Only")
    ws4.sheet_view.showGridLines = False

    ws4.merge_cells("A1:E1")
    c = ws4["A1"]
    c.value = "DISCREPANCIES ONLY \u2014 \u274c MISMATCH rows (Qty fields)"
    c.font = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    c.fill = fill(C["ORANGE"]); c.alignment = center(); ws4.row_dimensions[1].height = 22

    for ci, h in enumerate(det_hdrs, 1):
        wc(ws4, 2, ci, h, bg="37474F", fnt=hfont(), aln=center())
    ws4.row_dimensions[2].height = 16

    ri4 = 3
    for po in all_pos:
        df_f   = compare_po_fields(xl_df, pdf_data, po)
        if df_f.empty: continue
        issues = df_f[df_f["Status"] != "✅ MATCH"]
        if issues.empty: continue
        for _, row in issues.iterrows():
            alt = C["LGRAY"] if ri4 % 2 == 0 else C["WHITE"]
            ws4.row_dimensions[ri4].height = 14
            for ci, v in enumerate([po, row["Field"], row["Infor Value"], row["SAP Value"], row["Status"]], 1):
                if ci == 5:
                    wc(ws4, ri4, ci, v, bg=C["RED"],
                       fnt=cfont(bold=True, color="FFFFFF", sz=9), aln=center())
                elif ci in (3, 4):
                    wc(ws4, ri4, ci, v, bg="FFEBEE",
                       fnt=cfont(bold=True, color=C["RED"], sz=9), aln=center())
                elif ci == 1:
                    wc(ws4, ri4, ci, v, bg=alt, fnt=cfont(bold=True, sz=9), aln=left())
                else:
                    wc(ws4, ri4, ci, v, bg=alt, fnt=cfont(sz=9),
                       aln=left() if ci == 2 else center())
            ri4 += 1

    if ri4 == 3:
        c = ws4.cell(3, 1, "\u2705 No discrepancies found \u2014 all qty fields match!")
        c.font = Font(name="Arial", bold=True, color=C["GREEN"], size=11)

    for ci, w in enumerate([16, 24, 16, 16, 14], 1):
        ws4.column_dimensions[get_column_letter(ci)].width = w
    ws4.freeze_panes = "A3"

    # ─────────────────────────────────────────────────────────────────────────
    # Sheet 5 — Raw Infor Data
    # ─────────────────────────────────────────────────────────────────────────
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

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf


# ── Streamlit UI ──────────────────────────────────────────────────────────────
st.markdown(
    "<style>.stApp{font-family:Arial,sans-serif}.block-container{padding-top:1.2rem}</style>",
    unsafe_allow_html=True,
)

st.title("\U0001f4e6 Infor vs SAP Carton \u2014 Field-by-Field PO Comparator")
st.caption(
    "**Compared:** Total Qty (Pairs) · Qty per Size  |  "
    "**Info only (not compared):** Article · Model  |  "
    "**Size detail columns:** UK/US Size · XL Line · CTN Range · CTNs · Qty/CTN"
)

col1, col2 = st.columns(2)
with col1:
    st.subheader("\U0001f4ca Infor File \u2014 Order List (Excel)")
    xl_file = st.file_uploader("Upload Excel (.xlsx / .xls)", type=["xlsx", "xls"])
with col2:
    st.subheader("\U0001f4c4 SAP Carton Form (PDF)")
    pdf_file = st.file_uploader("Upload PDF (multi-page)", type=["pdf"])

if xl_file and pdf_file:
    with st.spinner("Processing & matching data..."):
        xl_df    = load_excel(xl_file)
        pdf_data = parse_pdf(pdf_file, filename=pdf_file.name)
        xl_pos     = xl_df["Order #"].str.strip().unique().tolist()
        pdf_pos    = list(pdf_data.keys())
        all_pos    = sorted(set(xl_pos + pdf_pos))
        match_pos  = [p for p in pdf_pos if p in xl_pos]
        xl_pos_set = set(xl_pos)

    # build summary
    sum_rows = []
    for po in all_pos:
        df_f    = compare_po_fields(xl_df, pdf_data, po)
        total_f = len(df_f)
        n_m     = (df_f["Status"] == "✅ MATCH").sum()    if not df_f.empty else 0
        n_mm    = (df_f["Status"] == "❌ MISMATCH").sum() if not df_f.empty else 0
        sap_src = pdf_data.get(po, {}).get("header", {}).get("Source File", pdf_file.name)

        if po not in pdf_data:     result = "⚠️ NO SAP DATA"
        elif po not in xl_pos_set: result = "⚠️ NO INFOR DATA"
        elif n_mm == 0:            result = "✅ ALL OK"
        else:                      result = f"❌ {n_mm} ISSUE(S)"

        sum_rows.append({
            "PO Number":    po,
            "Infor File":   xl_file.name,
            "SAP File(s)":  sap_src,
            "Total Fields": total_f,
            "✅ Match":     n_m,
            "❌ Mismatch":  n_mm,
            "Result":       result,
        })
    sum_df = pd.DataFrame(sum_rows)

    # KPIs
    k1,k2,k3,k4,k5,k6 = st.columns(6)
    k1.metric("Matched POs",             len(match_pos))
    k2.metric("✅ All OK",                (sum_df["Result"] == "✅ ALL OK").sum())
    k3.metric("❌ POs w/ Issues",         (sum_df["❌ Mismatch"] > 0).sum())
    k4.metric("Total Field Mismatches",   int(sum_df["❌ Mismatch"].sum()))
    k5.metric("SAP only",                len([p for p in pdf_pos if p not in xl_pos_set]))
    k6.metric("Infor only",              len([p for p in xl_pos  if p not in pdf_pos]))

    st.divider()

    # ── Summary Table ──
    st.subheader("📋 PO Compare Summary — adidas Infor vs SAP Carton")

    st.info(
        "ℹ️ **Fields compared:** Total Qty (Pairs) + Qty per Size.  "
        "Article & Model are **not compared** — shown as reference info in Size Detail tab.",
        icon=None,
    )

    def _cr(val):
        if "ALL OK" in str(val): return "background-color:#c8f7c5;color:#1B5E20;font-weight:bold"
        if "ISSUE"  in str(val): return "background-color:#ffcdd2;color:#B71C1C;font-weight:bold"
        if "NO"     in str(val): return "background-color:#ffe0b2;color:#E65100;font-weight:bold"
        return ""
    def _cmm(val):
        if isinstance(val, (int, float)) and val > 0:
            return "color:#B71C1C;font-weight:bold"
        return "color:#1B5E20;font-weight:bold"

    st.dataframe(
        sum_df.style.map(_cr, subset=["Result"]).map(_cmm, subset=["❌ Mismatch"]),
        use_container_width=True, hide_index=True,
        height=min(600, 80 + len(sum_df) * 36),
    )

    st.divider()
    tab1, tab2, tab3 = st.tabs([
        "🔍 Field-by-Field Detail",
        "📐 Size Detail (All Columns)",
        "❌ Discrepancies Only",
    ])

    # ── TAB 1: Field-by-Field ──────────────────────────────────────────────
    with tab1:
        st.caption(
            "**Compared fields:** `Total Qty (Pairs)` + `Qty Size X` per UK Size.  "
            "Article & Model are **not shown here** — see Size Detail tab for reference."
        )
        det_rows = []
        for po in all_pos:
            df_f = compare_po_fields(xl_df, pdf_data, po)
            if df_f.empty: continue
            df_f.insert(0, "PO Number", po)
            det_rows.append(df_f)

        if det_rows:
            det_df = pd.concat(det_rows, ignore_index=True)

            def _cs(val):
                if val == "✅ MATCH":    return "background-color:#c8f7c5;color:#1B5E20;font-weight:bold"
                if val == "❌ MISMATCH": return "background-color:#ffcdd2;color:#B71C1C;font-weight:bold"
                return ""

            def _cv(row):
                base = [""] * len(row)
                idx = list(row.index)
                if row.get("Status") == "❌ MISMATCH":
                    for f in ["Infor Value", "SAP Value"]:
                        if f in idx:
                            base[idx.index(f)] = "background-color:#ffebee;color:#B71C1C;font-weight:bold"
                return base

            st.dataframe(
                det_df.style.map(_cs, subset=["Status"]).apply(_cv, axis=1),
                use_container_width=True, hide_index=True,
                height=min(700, 80 + len(det_df) * 34),
            )

    # ── TAB 2: Size Detail ─────────────────────────────────────────────────
    with tab2:
        st.caption(
            "**All columns per size row.**  "
            "Article & Model shown as 🔵 info reference (not compared).  "
            "XL Line (Infor-only) · CTN Range · CTNs · Qty/CTN (SAP-only) shown as extra context columns."
        )
        sd_rows = []
        for po in all_pos:
            info = po_info(xl_df, pdf_data, po)
            sd   = compare_po_size_detail(xl_df, pdf_data, po)
            if sd.empty: continue
            sd.insert(0, "PO Number",       po)
            sd.insert(1, "Article (Infor)", info["Article (Infor)"])
            sd.insert(2, "Article (SAP)",   info["Article (SAP)"])
            sd.insert(3, "Model (Infor)",   info["Model (Infor)"])
            sd.insert(4, "Model (SAP)",     info["Model (SAP)"])
            sd_rows.append(sd)

        if sd_rows:
            sd_df = pd.concat(sd_rows, ignore_index=True)

            INFO_COLS = ["Article (Infor)", "Article (SAP)", "Model (Infor)", "Model (SAP)"]

            def _sst(val):
                if val == "✅ MATCH":    return "background-color:#c8f7c5;color:#1B5E20;font-weight:bold"
                if val == "❌ MISMATCH": return "background-color:#ffcdd2;color:#B71C1C;font-weight:bold"
                if "ONLY" in str(val):  return "background-color:#ffe0b2;color:#E65100;font-weight:bold"
                return ""

            def _sdiff(val):
                try:
                    if int(val) != 0: return "color:#B71C1C;font-weight:bold"
                except: pass
                return ""

            def _info_col(val):
                return "background-color:#E3F2FD;color:#0D47A1"

            present_info = [c for c in INFO_COLS if c in sd_df.columns]
            styled = (
                sd_df.style
                .map(_sst,      subset=["Status"])
                .map(_sdiff,    subset=["Diff"])
                .map(_info_col, subset=present_info)
            )
            st.dataframe(
                styled,
                use_container_width=True, hide_index=True,
                height=min(700, 80 + len(sd_df) * 34),
            )
        else:
            st.info("No size data available.")

    # ── TAB 3: Discrepancies Only ──────────────────────────────────────────
    with tab3:
        disc_rows = []
        for po in all_pos:
            df_f = compare_po_fields(xl_df, pdf_data, po)
            if df_f.empty: continue
            iss = df_f[df_f["Status"] != "✅ MATCH"].copy()
            if iss.empty: continue
            iss.insert(0, "PO Number", po)
            disc_rows.append(iss)

        if disc_rows:
            disc_df = pd.concat(disc_rows, ignore_index=True)

            def _ds(row):
                base = [""] * len(row)
                idx = list(row.index)
                for f in ["Infor Value", "SAP Value"]:
                    if f in idx:
                        base[idx.index(f)] = "background-color:#ffebee;color:#B71C1C;font-weight:bold"
                if "Status" in idx:
                    base[idx.index("Status")] = "background-color:#ffcdd2;color:#B71C1C;font-weight:bold"
                return base

            st.dataframe(
                disc_df.style.apply(_ds, axis=1),
                use_container_width=True, hide_index=True,
                height=min(500, 80 + len(disc_df) * 34),
            )
        else:
            st.success("✅ No discrepancies found — all matched POs are 100% OK!")

    # ── Download ──
    st.divider()
    st.subheader("⬇️ Download Excel Report")
    with st.spinner("Building Excel report..."):
        report_buf = build_report(xl_df, pdf_data, all_pos, xl_file.name, pdf_file.name)
    fname = f"InforVsSAP_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button(
        "📥 Download Excel Report",
        data=report_buf, file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, type="primary",
    )

    pdf_only = [p for p in pdf_pos if p not in xl_pos_set]
    xl_only  = [p for p in xl_pos  if p not in pdf_pos]
    if pdf_only:
        st.warning(f"⚠️ {len(pdf_only)} PO in SAP only (not in Infor): {', '.join(pdf_only)}")
    if xl_only:
        with st.expander(f"ℹ️ {len(xl_only)} PO in Infor only (not in SAP PDF)"):
            st.write(xl_only)

elif xl_file and not pdf_file:
    st.info("📄 Please upload the SAP Carton Form PDF.")
elif pdf_file and not xl_file:
    st.info("📊 Please upload the Infor Order List Excel.")
else:
    st.info("👆 Upload both files above to start.")
    with st.expander("ℹ️ How it works"):
        st.markdown("""
        **Match key**: `Order #` (Infor) = `Cust.PO` (SAP) — matched per PO number

        **Tab 1 — Field-by-Field Detail** `Field | Infor Value | SAP Value | Status`
        | Field | Compared? |
        |---|---|
        | Total Qty (Pairs) | ✅ Yes — ✅/❌ status |
        | Qty Size X (per UK Size) | ✅ Yes — ✅/❌ status |
        | Article | ❌ No — info only in Size Detail |
        | Model | ❌ No — info only in Size Detail |

        **Tab 2 — Size Detail** — one row per UK Size, all columns:
        `PO | Article (Infor) | Article (SAP) | Model (Infor) | Model (SAP) | UK Size | US Size | XL Line | Infor Qty | CTN Range | CTNs | Qty/CTN | SAP Qty | Diff | Status`
        - 🔵 Article & Model = info reference only (light blue)
        - XL Line = Infor-only column
        - CTN Range, CTNs, Qty/CTN = SAP-only columns

        **Excel report sheets:**
        1. `PO Compare Summary` · 2. `PO Compare Detail` · 3. `Size Detail` · 4. `Discrepancies Only` · 5. `Raw Infor Data`
        """)
