import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

st.set_page_config(
    page_title="Infor vs SAP Comparator",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}

.stApp { background-color: #F0F2F6; }
.block-container { padding: 1.5rem 2rem 3rem 2rem; max-width: 1400px; }

/* Header */
.app-header {
    background: linear-gradient(135deg, #0A1628 0%, #1A3A5C 100%);
    border-radius: 12px;
    padding: 24px 32px;
    margin-bottom: 24px;
    border-left: 5px solid #00D4AA;
}
.app-header h1 {
    color: #FFFFFF;
    font-size: 1.6rem;
    font-weight: 700;
    margin: 0 0 4px 0;
    font-family: 'IBM Plex Sans', sans-serif;
}
.app-header p {
    color: #94A3B8;
    font-size: 0.85rem;
    margin: 0;
}

/* KPI Cards */
.kpi-grid { display: grid; grid-template-columns: repeat(6, 1fr); gap: 12px; margin: 20px 0; }
.kpi-card {
    background: white;
    border-radius: 10px;
    padding: 16px;
    text-align: center;
    border-top: 4px solid #E2E8F0;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
}
.kpi-card.ok   { border-top-color: #10B981; }
.kpi-card.bad  { border-top-color: #EF4444; }
.kpi-card.warn { border-top-color: #F59E0B; }
.kpi-card.blue { border-top-color: #3B82F6; }
.kpi-number { font-size: 2rem; font-weight: 700; color: #1E293B; line-height: 1; margin-bottom: 4px; font-family: 'IBM Plex Mono', monospace; }
.kpi-label  { font-size: 0.72rem; color: #64748B; text-transform: uppercase; letter-spacing: 0.05em; font-weight: 600; }
.kpi-card.ok   .kpi-number { color: #10B981; }
.kpi-card.bad  .kpi-number { color: #EF4444; }
.kpi-card.warn .kpi-number { color: #F59E0B; }

/* PO Card */
.po-card {
    background: white;
    border-radius: 12px;
    margin-bottom: 16px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    overflow: hidden;
}
.po-card-header {
    padding: 14px 20px;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.po-card-header.ok   { background: linear-gradient(90deg, #F0FDF4, #DCFCE7); border-left: 5px solid #10B981; }
.po-card-header.bad  { background: linear-gradient(90deg, #FFF1F2, #FFE4E6); border-left: 5px solid #EF4444; }
.po-card-header.warn { background: linear-gradient(90deg, #FFFBEB, #FEF3C7); border-left: 5px solid #F59E0B; }
.po-number { font-size: 1.1rem; font-weight: 700; color: #1E293B; font-family: 'IBM Plex Mono', monospace; }
.po-badge {
    padding: 4px 14px;
    border-radius: 999px;
    font-size: 0.78rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}
.po-badge.ok   { background: #10B981; color: white; }
.po-badge.bad  { background: #EF4444; color: white; }
.po-badge.warn { background: #F59E0B; color: white; }

/* Compare Table */
.cmp-table { width: 100%; border-collapse: collapse; font-size: 0.82rem; }
.cmp-table th {
    background: #1E293B;
    color: white;
    padding: 10px 14px;
    text-align: center;
    font-weight: 600;
    font-size: 0.75rem;
    text-transform: uppercase;
    letter-spacing: 0.06em;
}
.cmp-table th.left { text-align: left; }
.cmp-table td {
    padding: 9px 14px;
    border-bottom: 1px solid #F1F5F9;
    text-align: center;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.82rem;
}
.cmp-table td.label {
    text-align: left;
    font-family: 'IBM Plex Sans', sans-serif;
    font-weight: 500;
    color: #374151;
}
.cmp-table tr:hover td { background: #F8FAFC; }
.cmp-table tr.match td { }
.cmp-table tr.mismatch td.val { background: #FEF2F2; color: #B91C1C; font-weight: 700; }
.cmp-table tr.only-infor td { background: #FFFBEB; }
.cmp-table tr.only-sap td   { background: #EFF6FF; }

.badge-match    { background:#D1FAE5; color:#065F46; padding:3px 10px; border-radius:999px; font-weight:700; font-size:0.72rem; font-family:'IBM Plex Sans',sans-serif; }
.badge-mismatch { background:#FEE2E2; color:#991B1B; padding:3px 10px; border-radius:999px; font-weight:700; font-size:0.72rem; font-family:'IBM Plex Sans',sans-serif; }
.badge-only-i   { background:#FEF3C7; color:#92400E; padding:3px 10px; border-radius:999px; font-weight:700; font-size:0.72rem; font-family:'IBM Plex Sans',sans-serif; }
.badge-only-s   { background:#DBEAFE; color:#1E40AF; padding:3px 10px; border-radius:999px; font-weight:700; font-size:0.72rem; font-family:'IBM Plex Sans',sans-serif; }

/* Info row in size table */
.info-cell { color: #3B82F6; font-size: 0.78rem; font-style: italic; font-family:'IBM Plex Sans',sans-serif; }

/* Progress bar */
.match-bar-wrap { background: #E2E8F0; border-radius: 999px; height: 8px; overflow: hidden; margin: 8px 0 4px 0; }
.match-bar-fill { height: 100%; border-radius: 999px; transition: width 0.5s; }

/* Section label */
.section-label {
    font-size: 0.7rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    color: #94A3B8;
    margin: 20px 0 8px 0;
}

/* Upload zone */
.upload-card {
    background: white;
    border-radius: 12px;
    padding: 20px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    border: 2px dashed #CBD5E1;
}

/* Tab style override */
.stTabs [data-baseweb="tab-list"] {
    background: white;
    border-radius: 10px 10px 0 0;
    padding: 4px 8px 0;
    gap: 4px;
    border-bottom: 2px solid #E2E8F0;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px 8px 0 0;
    font-weight: 600;
    font-size: 0.85rem;
    padding: 10px 20px;
    color: #64748B;
}
.stTabs [aria-selected="true"] {
    background: #0A1628 !important;
    color: white !important;
}
</style>
""", unsafe_allow_html=True)

# ── Colors & Style Helpers (Excel) ────────────────────────────────────────────
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
        "Pack Mode":   str(r.get("VAS/SHAS L15 – Packing Mode", "")),
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
            if not po: continue
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
                r"(\d+[-–]\d+|\d+)\s+(\d+)\s+(\d+)\s+(\d+)"
                r"\s+([\d]+[-K]?[-]?)\s+([\d]+[-K]?[-]?)\s+[\d.]+"
            )
            rows = [
                {"CTN Range": m.group(1), "CTNs": int(m.group(2)),
                 "Qty/CTN": int(m.group(3)), "SAP Qty": int(m.group(4)),
                 "UK Size": m.group(5).strip(), "US Size": m.group(6).strip()}
                for m in pat.finditer(text)
            ]
            result[po] = {
                "header": hdr,
                "sizes":  pd.DataFrame(rows) if rows else pd.DataFrame(),
            }
    return result

# ── Comparison Logic ──────────────────────────────────────────────────────────
def compare_po_fields(xl_df, pdf_data, po):
    xs = xl_sizes_for_po(xl_df, po)
    pg = pdf_data.get(po, {})
    ph = pg.get("header", {})
    ps = pg.get("sizes", pd.DataFrame())
    rows = []
    infor_total = int(xs["Infor Qty"].sum()) if not xs.empty else 0
    sap_total   = int(ph.get("Total Pairs", 0) or 0)
    rows.append({"Field": "Total Qty (Pairs)", "Infor Value": infor_total,
                 "SAP Value": sap_total,
                 "Status": "✅ MATCH" if infor_total == sap_total else "❌ MISMATCH"})
    if not xs.empty and ps is not None and not ps.empty:
        mg = pd.merge(xs[["UK Size","Infor Qty"]], ps[["UK Size","SAP Qty"]], on="UK Size", how="outer")
    elif not xs.empty:
        mg = xs[["UK Size","Infor Qty"]].copy(); mg["SAP Qty"] = 0
    elif ps is not None and not ps.empty:
        mg = ps[["UK Size","SAP Qty"]].copy(); mg["Infor Qty"] = 0
    else:
        mg = pd.DataFrame(columns=["UK Size","Infor Qty","SAP Qty"])
    if not mg.empty:
        mg["Infor Qty"] = mg["Infor Qty"].fillna(0).astype(int)
        mg["SAP Qty"]   = mg["SAP Qty"].fillna(0).astype(int)
        mg = mg.sort_values("UK Size").reset_index(drop=True)
        for _, row in mg.iterrows():
            iv, sv = int(row["Infor Qty"]), int(row["SAP Qty"])
            rows.append({"Field": f"Qty Size {row['UK Size']}", "Infor Value": iv,
                         "SAP Value": sv, "Status": "✅ MATCH" if iv==sv else "❌ MISMATCH"})
    return pd.DataFrame(rows)

def compare_po_size_detail(xl_df, pdf_data, po):
    xs = xl_sizes_for_po(xl_df, po)
    pg = pdf_data.get(po, {})
    ps = pg.get("sizes", pd.DataFrame())
    COLS = ["UK Size","US Size","XL Line","Infor Qty","CTN Range","CTNs","Qty/CTN","SAP Qty","Diff","Status"]
    if not xs.empty and ps is not None and not ps.empty:
        xl_c  = xs[["UK Size","US Size","Infor Qty","XL Line"]]
        sap_c = ps[["UK Size","US Size","SAP Qty","CTN Range","CTNs","Qty/CTN"]].rename(columns={"US Size":"_sap_us"})
        mg = pd.merge(xl_c, sap_c, on="UK Size", how="outer")
        mg["US Size"] = mg["US Size"].combine_first(mg["_sap_us"])
        mg = mg.drop(columns=["_sap_us"], errors="ignore")
    elif not xs.empty:
        mg = xs[["UK Size","US Size","Infor Qty","XL Line"]].copy()
        mg["SAP Qty"] = 0
        for col in ["CTN Range","CTNs","Qty/CTN"]: mg[col] = ""
    elif ps is not None and not ps.empty:
        mg = ps[["UK Size","US Size","SAP Qty","CTN Range","CTNs","Qty/CTN"]].copy()
        mg["Infor Qty"] = 0; mg["XL Line"] = ""
    else:
        return pd.DataFrame(columns=COLS)
    mg["Infor Qty"] = mg["Infor Qty"].fillna(0).astype(int)
    mg["SAP Qty"]   = mg["SAP Qty"].fillna(0).astype(int)
    mg["Diff"]      = mg["Infor Qty"] - mg["SAP Qty"]
    def _st(row):
        no_sap   = str(row.get("CTN Range","")).strip() in ("","nan","None","NaN")
        no_infor = str(row.get("XL Line","")).strip()   in ("","nan","None","NaN")
        if no_sap:   return "ONLY IN INFOR"
        if no_infor: return "ONLY IN SAP"
        return "MATCH" if row["Diff"]==0 else "MISMATCH"
    mg["Status"] = mg.apply(_st, axis=1)
    for col in COLS:
        if col not in mg.columns: mg[col] = ""
    return mg[COLS].sort_values("UK Size").reset_index(drop=True)

def po_info(xl_df, pdf_data, po):
    xh = xl_hdr_for_po(xl_df, po)
    ph = pdf_data.get(po, {}).get("header", {})
    return {
        "Article (Infor)": xh.get("Article",""),
        "Article (SAP)":   ph.get("Article",""),
        "Model (Infor)":   xh.get("Model",""),
        "Model (SAP)":     ph.get("Model",""),
        "Total Pairs (SAP)": ph.get("Total Pairs",""),
        "Pack Mode (SAP)":   ph.get("Pack Mode",""),
        "Total CTNs (SAP)":  ph.get("Total CTNs",""),
    }

# ── HTML Rendering Helpers ────────────────────────────────────────────────────
def status_badge(status):
    if status == "MATCH":         return '<span class="badge-match">✓ MATCH</span>'
    if status == "MISMATCH":      return '<span class="badge-mismatch">✗ MISMATCH</span>'
    if status == "ONLY IN INFOR": return '<span class="badge-only-i">⬛ INFOR ONLY</span>'
    if status == "ONLY IN SAP":   return '<span class="badge-only-s">⬛ SAP ONLY</span>'
    return status

def render_po_header(po, n_match, n_mismatch, n_only, in_pdf, in_xl):
    if not in_pdf:   cls, badge_cls, badge_txt = "warn", "warn", "⚠ NO SAP DATA"
    elif not in_xl:  cls, badge_cls, badge_txt = "warn", "warn", "⚠ NO INFOR DATA"
    elif n_mismatch == 0 and n_only == 0: cls, badge_cls, badge_txt = "ok", "ok", "✓ ALL MATCH"
    else:            cls, badge_cls, badge_txt = "bad", "bad", f"✗ {n_mismatch + n_only} ISSUE(S)"
    total = n_match + n_mismatch + n_only
    pct   = int(n_match / total * 100) if total > 0 else 0
    bar_color = "#10B981" if pct==100 else "#EF4444" if pct < 70 else "#F59E0B"
    return cls, f"""
    <div class="po-card-header {cls}">
        <div>
            <div class="po-number">PO&nbsp;{po}</div>
            <div style="font-size:0.78rem;color:#64748B;margin-top:2px;">
                {n_match} match · {n_mismatch} mismatch · {n_only} only-in-one-system
            </div>
            <div class="match-bar-wrap" style="width:200px">
                <div class="match-bar-fill" style="width:{pct}%;background:{bar_color}"></div>
            </div>
            <div style="font-size:0.7rem;color:{bar_color};font-weight:700;">{pct}% matched</div>
        </div>
        <span class="po-badge {badge_cls}">{badge_txt}</span>
    </div>
    """

def render_size_table(sd, info):
    rows_html = ""
    for _, row in sd.iterrows():
        st   = row.get("Status","")
        diff = row.get("Diff", 0)
        try: diff_i = int(diff)
        except: diff_i = 0

        if st == "MATCH":
            tr_cls   = "match"
            val_cls  = ""
            diff_cell = f'<td style="color:#10B981;font-weight:700;">0</td>'
        elif st == "MISMATCH":
            tr_cls   = "mismatch"
            val_cls  = " val"
            diff_cell = f'<td style="color:#EF4444;font-weight:700;">{diff_i:+d}</td>'
        elif st == "ONLY IN INFOR":
            tr_cls   = "only-infor"
            val_cls  = ""
            diff_cell= f'<td style="color:#F59E0B;font-weight:700;">–</td>'
        else:
            tr_cls   = "only-sap"
            val_cls  = ""
            diff_cell= f'<td style="color:#3B82F6;font-weight:700;">–</td>'

        infor_qty = int(row.get("Infor Qty", 0)) if str(row.get("Infor Qty","")).strip() not in ("","nan") else "–"
        sap_qty   = int(row.get("SAP Qty", 0))   if str(row.get("SAP Qty","")).strip()   not in ("","nan") else "–"
        ctn_range = row.get("CTN Range","") or "–"
        ctns      = row.get("CTNs","")      or "–"
        qpc       = row.get("Qty/CTN","")   or "–"
        xl_line   = row.get("XL Line","")   or "–"

        rows_html += f"""
        <tr class="{tr_cls}">
            <td class="label">{row.get('UK Size','')}</td>
            <td>{row.get('US Size','')}</td>
            <td style="color:#6B7280;font-size:0.75rem;">{xl_line}</td>
            <td class="{val_cls}" style="font-weight:600;">{infor_qty}</td>
            <td class="{val_cls}" style="font-weight:600;">{sap_qty}</td>
            {diff_cell}
            <td style="font-size:0.75rem;color:#6B7280;">{ctn_range}</td>
            <td style="font-size:0.75rem;color:#6B7280;">{ctns}</td>
            <td style="font-size:0.75rem;color:#6B7280;">{qpc}</td>
            <td style="text-align:center;">{status_badge(st)}</td>
        </tr>
        """

    return f"""
    <table class="cmp-table">
        <thead>
            <tr>
                <th class="left">UK Size</th>
                <th>US Size</th>
                <th>XL Line</th>
                <th style="background:#1A3A5C;">Infor Qty</th>
                <th style="background:#1A3A5C;">SAP Qty</th>
                <th style="background:#DC2626;">Diff</th>
                <th style="background:#374151;">CTN Range</th>
                <th style="background:#374151;">CTNs</th>
                <th style="background:#374151;">Qty/CTN</th>
                <th style="background:#374151;">Status</th>
            </tr>
        </thead>
        <tbody>
            {rows_html}
        </tbody>
    </table>
    """

def render_info_strip(info):
    def pair(label, v1, v2):
        match_icon = "✓" if v1 and v2 and v1==v2 else ("–" if not v1 or not v2 else "≠")
        color = "#10B981" if match_icon=="✓" else "#94A3B8" if match_icon=="–" else "#F59E0B"
        return f"""
        <div style="flex:1;min-width:180px;background:#F8FAFC;border-radius:8px;padding:10px 14px;border:1px solid #E2E8F0;">
            <div style="font-size:0.65rem;text-transform:uppercase;letter-spacing:0.08em;color:#94A3B8;font-weight:700;margin-bottom:4px;">{label}</div>
            <div style="font-size:0.78rem;color:#1E293B;font-weight:600;">📊 {v1 or '<em style="color:#CBD5E1">–</em>'}</div>
            <div style="font-size:0.78rem;color:#1E293B;font-weight:600;margin-top:2px;">📄 {v2 or '<em style="color:#CBD5E1">–</em>'}</div>
            <div style="font-size:0.72rem;color:{color};font-weight:700;margin-top:4px;">{match_icon}</div>
        </div>
        """
    return f"""
    <div style="display:flex;gap:10px;flex-wrap:wrap;padding:14px 20px;background:#FAFBFC;border-top:1px solid #F1F5F9;">
        {pair("Article", info.get("Article (Infor)",""), info.get("Article (SAP)",""))}
        {pair("Model",   info.get("Model (Infor)",""),   info.get("Model (SAP)",""))}
        <div style="flex:1;min-width:120px;background:#F8FAFC;border-radius:8px;padding:10px 14px;border:1px solid #E2E8F0;">
            <div style="font-size:0.65rem;text-transform:uppercase;letter-spacing:0.08em;color:#94A3B8;font-weight:700;margin-bottom:4px;">SAP Pack Mode</div>
            <div style="font-size:0.88rem;color:#1E293B;font-weight:700;">{info.get("Pack Mode (SAP)","–") or "–"}</div>
        </div>
        <div style="flex:1;min-width:120px;background:#F8FAFC;border-radius:8px;padding:10px 14px;border:1px solid #E2E8F0;">
            <div style="font-size:0.65rem;text-transform:uppercase;letter-spacing:0.08em;color:#94A3B8;font-weight:700;margin-bottom:4px;">SAP Total CTNs</div>
            <div style="font-size:0.88rem;color:#1E293B;font-weight:700;">{info.get("Total CTNs (SAP)","–") or "–"}</div>
        </div>
        <div style="flex:1;min-width:140px;background:#F8FAFC;border-radius:8px;padding:10px 14px;border:1px solid #E2E8F0;">
            <div style="font-size:0.65rem;text-transform:uppercase;letter-spacing:0.08em;color:#94A3B8;font-weight:700;margin-bottom:4px;">SAP Total Pairs</div>
            <div style="font-size:0.88rem;color:#1E293B;font-weight:700;">{info.get("Total Pairs (SAP)","–") or "–"}</div>
        </div>
    </div>
    """

# ── Legend ────────────────────────────────────────────────────────────────────
def render_legend():
    return """
    <div style="display:flex;gap:12px;flex-wrap:wrap;align-items:center;
                background:white;border-radius:8px;padding:10px 16px;
                border:1px solid #E2E8F0;margin-bottom:16px;font-size:0.78rem;">
        <strong style="color:#374151;margin-right:4px;">Legend:</strong>
        <span class="badge-match">✓ MATCH</span>
        <span style="color:#6B7280">Qty Infor = SAP</span>
        &nbsp;·&nbsp;
        <span class="badge-mismatch">✗ MISMATCH</span>
        <span style="color:#6B7280">Qty berbeda</span>
        &nbsp;·&nbsp;
        <span class="badge-only-i">⬛ INFOR ONLY</span>
        <span style="color:#6B7280">Ada di Infor, tidak di SAP</span>
        &nbsp;·&nbsp;
        <span class="badge-only-s">⬛ SAP ONLY</span>
        <span style="color:#6B7280">Ada di SAP, tidak di Infor</span>
    </div>
    """

# ── Excel Report Builder ──────────────────────────────────────────────────────
def build_report(xl_df, pdf_data, all_pos, xl_filename, pdf_filename):
    wb = Workbook()
    xl_pos_set = set(xl_df["Order #"].str.strip().tolist())
    ST_STYLE = {
        "✅ MATCH":      (C["GREEN"],  "FFFFFF"),
        "❌ MISMATCH":   (C["RED"],    "FFFFFF"),
        "ONLY IN INFOR": (C["ORANGE"], "FFFFFF"),
        "ONLY IN SAP":   (C["ORANGE"], "FFFFFF"),
    }

    # Sheet 1 — Summary
    ws1 = wb.active; ws1.title = "PO Compare Summary"
    ws1.sheet_view.showGridLines = False
    ws1.merge_cells("A1:G1")
    c = ws1["A1"]; c.value = "PO COMPARE SUMMARY — Infor vs SAP Carton"
    c.font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    c.fill = fill(C["DGRAY"]); c.alignment = center(); ws1.row_dimensions[1].height = 28
    ws1.merge_cells("A2:G2")
    c = ws1["A2"]
    c.value = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  Infor: {xl_filename}  |  SAP: {pdf_filename}"
    c.font = Font(name="Arial", size=9, color="616161")
    c.fill = fill(C["LGRAY"]); c.alignment = left(); ws1.row_dimensions[2].height = 14
    hdrs = ["PO Number","Infor File","SAP File(s)","Total Fields","✅ Match","❌ Mismatch","Result"]
    for ci, h in enumerate(hdrs, 1):
        wc(ws1, 4, ci, h, bg=C["DGRAY"], fnt=hfont(), aln=center())
    ws1.row_dimensions[4].height = 16
    for ri, po in enumerate(all_pos, 5):
        df_f = compare_po_fields(xl_df, pdf_data, po)
        total_f = len(df_f)
        n_m  = (df_f["Status"]=="✅ MATCH").sum()    if not df_f.empty else 0
        n_mm = (df_f["Status"]=="❌ MISMATCH").sum() if not df_f.empty else 0
        sap_src = pdf_data.get(po,{}).get("header",{}).get("Source File", pdf_filename)
        if po not in pdf_data:     r_txt,r_bg,r_fc = "⚠️ NO SAP DATA",   C["ORANGE"],"FFFFFF"
        elif po not in xl_pos_set: r_txt,r_bg,r_fc = "⚠️ NO INFOR DATA", C["ORANGE"],"FFFFFF"
        elif n_mm==0:              r_txt,r_bg,r_fc = "✅ ALL OK",          C["GREEN"], "FFFFFF"
        else:                      r_txt,r_bg,r_fc = f"❌ {n_mm} ISSUE(S)",C["RED"],  "FFFFFF"
        alt = C["LGRAY"] if ri%2==0 else C["WHITE"]
        ws1.row_dimensions[ri].height = 15
        for ci, v in enumerate([po,xl_filename,sap_src,total_f,n_m,n_mm,r_txt], 1):
            if ci==7: wc(ws1, ri, ci, v, bg=r_bg, fnt=hfont(color=r_fc), aln=center())
            elif ci==5: wc(ws1, ri, ci, v, bg=alt, fnt=cfont(bold=True,color=C["DARKGREEN"],sz=9), aln=center())
            elif ci==6 and n_mm>0: wc(ws1, ri, ci, v, bg=alt, fnt=cfont(bold=True,color=C["RED"],sz=9), aln=center())
            else: wc(ws1, ri, ci, v, bg=alt, fnt=cfont(sz=9), aln=left() if ci<=3 else center())
    for ci, w in enumerate([16,36,36,13,10,13,18], 1):
        ws1.column_dimensions[get_column_letter(ci)].width = w
    ws1.freeze_panes = "A5"

    # Sheet 2 — Detail
    ws2 = wb.create_sheet("PO Compare Detail")
    ws2.sheet_view.showGridLines = False
    ws2.merge_cells("A1:E1")
    c = ws2["A1"]; c.value = "PO COMPARE DETAIL — Field by Field"
    c.font = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    c.fill = fill(C["DGRAY"]); c.alignment = center(); ws2.row_dimensions[1].height = 24
    det_hdrs = ["PO Number","Field","Infor Value","SAP Value","Status"]
    for ci, h in enumerate(det_hdrs, 1):
        wc(ws2, 2, ci, h, bg="37474F", fnt=hfont(), aln=center())
    ws2.row_dimensions[2].height = 16
    ri2 = 3
    for po in all_pos:
        df_f = compare_po_fields(xl_df, pdf_data, po)
        if df_f.empty: continue
        for _, row in df_f.iterrows():
            st = row["Status"]; is_m = (st=="✅ MATCH")
            alt = C["LGRAY"] if ri2%2==0 else C["WHITE"]
            ws2.row_dimensions[ri2].height = 14
            for ci, v in enumerate([po,row["Field"],row["Infor Value"],row["SAP Value"],st], 1):
                if ci==5: wc(ws2, ri2, ci, v, bg=C["GREEN"] if is_m else C["RED"], fnt=cfont(bold=True,color="FFFFFF",sz=9), aln=center())
                elif ci in (3,4) and not is_m: wc(ws2, ri2, ci, v, bg="FFEBEE", fnt=cfont(bold=True,color=C["RED"],sz=9), aln=center())
                elif ci==1: wc(ws2, ri2, ci, v, bg=alt, fnt=cfont(bold=True,sz=9), aln=left())
                else: wc(ws2, ri2, ci, v, bg=alt, fnt=cfont(sz=9), aln=left() if ci==2 else center())
            ri2 += 1
    for ci, w in enumerate([16,24,16,16,14], 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w
    ws2.freeze_panes = "A3"

    # Sheet 3 — Size Detail
    ws3 = wb.create_sheet("Size Detail")
    ws3.sheet_view.showGridLines = False
    ws3.merge_cells("A1:M1")
    c = ws3["A1"]; c.value = "SIZE DETAIL — All POs"
    c.font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    c.fill = fill(C["DGRAY"]); c.alignment = left(); ws3.row_dimensions[1].height = 20
    sd_hdrs = ["PO Number","Article (Infor)","Article (SAP)","Model (Infor)","Model (SAP)",
               "UK Size","US Size","XL Line","Infor Qty","CTN Range","CTNs","Qty/CTN","SAP Qty","Diff","Status"]
    for ci, h in enumerate(sd_hdrs, 1):
        wc(ws3, 2, ci, h, bg="37474F", fnt=hfont(), aln=center())
    ws3.row_dimensions[2].height = 16
    ri3 = 3
    for po in all_pos:
        info = po_info(xl_df, pdf_data, po)
        sd   = compare_po_size_detail(xl_df, pdf_data, po)
        if sd.empty: continue
        for _, row in sd.iterrows():
            st = row.get("Status","")
            s_bg, s_fc = ST_STYLE.get("✅ MATCH" if st=="MATCH" else "❌ MISMATCH" if st=="MISMATCH" else "ONLY IN INFOR", (C["WHITE"],"000000"))
            alt = C["LGRAY"] if ri3%2==0 else C["WHITE"]
            ws3.row_dimensions[ri3].height = 14
            try: diff_i = int(row.get("Diff",0))
            except: diff_i = 0
            vals = [po, info["Article (Infor)"], info["Article (SAP)"],
                    info["Model (Infor)"], info["Model (SAP)"],
                    row.get("UK Size",""), row.get("US Size",""), row.get("XL Line",""),
                    int(row.get("Infor Qty",0)), row.get("CTN Range",""),
                    row.get("CTNs",""), row.get("Qty/CTN",""),
                    int(row.get("SAP Qty",0)), diff_i, st]
            for ci, v in enumerate(vals, 1):
                if ci==15: wc(ws3, ri3, ci, v, bg=s_bg, fnt=cfont(bold=True,color=s_fc,sz=9), aln=center())
                elif ci==14 and diff_i!=0: wc(ws3, ri3, ci, v, bg=C["RED"], fnt=hfont(sz=9), aln=center())
                elif ci in (2,3,4,5): wc(ws3, ri3, ci, v, bg=C["INFO_BG"], fnt=cfont(sz=9,color=C["INFO_FG"]), aln=left())
                else: wc(ws3, ri3, ci, v, bg=alt, fnt=cfont(sz=9), aln=left() if ci<=5 else center())
            ri3 += 1
    for ci, w in enumerate([14,12,12,22,22,9,9,9,10,13,8,9,10,8,16], 1):
        ws3.column_dimensions[get_column_letter(ci)].width = w
    ws3.freeze_panes = "A3"

    # Sheet 4 — Discrepancies
    ws4 = wb.create_sheet("Discrepancies Only")
    ws4.sheet_view.showGridLines = False
    ws4.merge_cells("A1:E1")
    c = ws4["A1"]; c.value = "DISCREPANCIES ONLY — ❌ MISMATCH rows"
    c.font = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    c.fill = fill(C["ORANGE"]); c.alignment = center(); ws4.row_dimensions[1].height = 22
    for ci, h in enumerate(det_hdrs, 1):
        wc(ws4, 2, ci, h, bg="37474F", fnt=hfont(), aln=center())
    ws4.row_dimensions[2].height = 16
    ri4 = 3
    for po in all_pos:
        df_f   = compare_po_fields(xl_df, pdf_data, po)
        if df_f.empty: continue
        issues = df_f[df_f["Status"]!="✅ MATCH"]
        if issues.empty: continue
        for _, row in issues.iterrows():
            alt = C["LGRAY"] if ri4%2==0 else C["WHITE"]
            ws4.row_dimensions[ri4].height = 14
            for ci, v in enumerate([po,row["Field"],row["Infor Value"],row["SAP Value"],row["Status"]], 1):
                if ci==5: wc(ws4, ri4, ci, v, bg=C["RED"], fnt=cfont(bold=True,color="FFFFFF",sz=9), aln=center())
                elif ci in (3,4): wc(ws4, ri4, ci, v, bg="FFEBEE", fnt=cfont(bold=True,color=C["RED"],sz=9), aln=center())
                elif ci==1: wc(ws4, ri4, ci, v, bg=alt, fnt=cfont(bold=True,sz=9), aln=left())
                else: wc(ws4, ri4, ci, v, bg=alt, fnt=cfont(sz=9), aln=left() if ci==2 else center())
            ri4 += 1
    if ri4==3:
        c = ws4.cell(3,1,"✅ No discrepancies found — all qty fields match!")
        c.font = Font(name="Arial", bold=True, color=C["GREEN"], size=11)
    for ci, w in enumerate([16,24,16,16,14], 1):
        ws4.column_dimensions[get_column_letter(ci)].width = w
    ws4.freeze_panes = "A3"

    # Sheet 5 — Raw
    ws_r = wb.create_sheet("Raw Infor Data")
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

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ═══════════════════════════════════════════════════════════════════════════════
# STREAMLIT UI
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="app-header">
    <h1>📦 Infor vs SAP — PO Comparator</h1>
    <p>Bandingkan data Order (Infor Excel) dengan Carton Form (SAP PDF) · per PO · per Ukuran · per Qty</p>
</div>
""", unsafe_allow_html=True)

# ── Upload ────────────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    st.markdown("**📊 Infor File — Order List (Excel)**")
    xl_file  = st.file_uploader("Upload Excel (.xlsx / .xls)", type=["xlsx","xls"], label_visibility="collapsed")
    if xl_file:  st.success(f"✓ {xl_file.name}")
with col2:
    st.markdown("**📄 SAP Carton Form (PDF)**")
    pdf_file = st.file_uploader("Upload PDF (multi-page)", type=["pdf"], label_visibility="collapsed")
    if pdf_file: st.success(f"✓ {pdf_file.name}")

if not (xl_file and pdf_file):
    st.markdown("""
    <div style="background:white;border-radius:12px;padding:28px 32px;margin-top:24px;
                border:1px solid #E2E8F0;text-align:center;color:#94A3B8;">
        <div style="font-size:2.5rem;margin-bottom:12px;">⬆️</div>
        <div style="font-size:1rem;font-weight:600;color:#374151;margin-bottom:8px;">Upload kedua file untuk memulai</div>
        <div style="font-size:0.82rem;">
            <b>Match key:</b> <code>Order #</code> (Infor) = <code>Cust.PO</code> (SAP)<br>
            <b>Yang dibandingkan:</b> Total Qty (Pairs) + Qty per UK Size<br>
            <b>Info saja (tidak dibandingkan):</b> Article · Model
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ── Process ───────────────────────────────────────────────────────────────────
with st.spinner("Memproses & mencocokkan data..."):
    xl_df    = load_excel(xl_file)
    pdf_data = parse_pdf(pdf_file, filename=pdf_file.name)
    xl_pos     = xl_df["Order #"].str.strip().unique().tolist()
    pdf_pos    = list(pdf_data.keys())
    all_pos    = sorted(set(xl_pos + pdf_pos))
    match_pos  = [p for p in pdf_pos if p in set(xl_pos)]
    xl_pos_set = set(xl_pos)

# Build summary data
sum_rows = []
for po in all_pos:
    df_f    = compare_po_fields(xl_df, pdf_data, po)
    n_m  = (df_f["Status"]=="✅ MATCH").sum()    if not df_f.empty else 0
    n_mm = (df_f["Status"]=="❌ MISMATCH").sum() if not df_f.empty else 0
    if po not in pdf_data:     result = "⚠️ NO SAP"
    elif po not in xl_pos_set: result = "⚠️ NO INFOR"
    elif n_mm==0:              result = "✅ ALL OK"
    else:                      result = f"❌ {n_mm} ISSUE(S)"
    sum_rows.append({"PO": po, "Match": n_m, "Mismatch": n_mm, "Result": result})
sum_df = pd.DataFrame(sum_rows)

# ── KPI ───────────────────────────────────────────────────────────────────────
n_ok     = (sum_df["Result"]=="✅ ALL OK").sum()
n_issues = (sum_df["Mismatch"]>0).sum()
n_total_mm = int(sum_df["Mismatch"].sum())
n_sap_only = len([p for p in pdf_pos if p not in xl_pos_set])
n_xl_only  = len([p for p in xl_pos  if p not in set(pdf_pos)])
match_rate = int(n_ok / len(all_pos) * 100) if all_pos else 0
bar_color  = "#10B981" if match_rate==100 else "#EF4444" if match_rate<70 else "#F59E0B"

st.markdown(f"""
<div class="kpi-grid">
    <div class="kpi-card blue">
        <div class="kpi-number">{len(match_pos)}</div>
        <div class="kpi-label">PO Matched</div>
    </div>
    <div class="kpi-card ok">
        <div class="kpi-number">{n_ok}</div>
        <div class="kpi-label">✅ All OK</div>
    </div>
    <div class="kpi-card bad">
        <div class="kpi-number">{n_issues}</div>
        <div class="kpi-label">❌ Ada Masalah</div>
    </div>
    <div class="kpi-card bad">
        <div class="kpi-number">{n_total_mm}</div>
        <div class="kpi-label">Total Mismatch</div>
    </div>
    <div class="kpi-card warn">
        <div class="kpi-number">{n_sap_only}</div>
        <div class="kpi-label">SAP Only</div>
    </div>
    <div class="kpi-card warn">
        <div class="kpi-number">{n_xl_only}</div>
        <div class="kpi-label">Infor Only</div>
    </div>
</div>

<div style="background:white;border-radius:10px;padding:14px 20px;margin-bottom:20px;
            border:1px solid #E2E8F0;display:flex;align-items:center;gap:16px;">
    <div style="font-size:0.8rem;font-weight:700;color:#374151;white-space:nowrap;">Overall Match Rate</div>
    <div style="flex:1;">
        <div class="match-bar-wrap">
            <div class="match-bar-fill" style="width:{match_rate}%;background:{bar_color}"></div>
        </div>
    </div>
    <div style="font-size:1.4rem;font-weight:700;color:{bar_color};font-family:'IBM Plex Mono',monospace;
                white-space:nowrap;">{match_rate}%</div>
    <div style="font-size:0.75rem;color:#94A3B8;">{n_ok} dari {len(all_pos)} PO</div>
</div>
""", unsafe_allow_html=True)

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs([
    f"📦 Per-PO Detail  ({len(all_pos)} PO)",
    f"❌ Discrepancies Only  ({n_total_mm})",
    "⬇️ Download Report",
])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — Per-PO Detail
# ═══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown(render_legend(), unsafe_allow_html=True)

    # Filter bar
    fcol1, fcol2 = st.columns([3,1])
    with fcol1:
        filter_po = st.text_input("🔍 Cari PO Number", placeholder="Ketik nomor PO...", label_visibility="collapsed")
    with fcol2:
        filter_status = st.selectbox("Filter", ["Semua","✅ OK Saja","❌ Ada Masalah","⚠️ Data Hilang"], label_visibility="collapsed")

    filtered_pos = all_pos
    if filter_po.strip():
        filtered_pos = [p for p in all_pos if filter_po.strip() in p]
    if filter_status == "✅ OK Saja":
        filtered_pos = [p for p in filtered_pos if sum_df[sum_df["PO"]==p]["Result"].values[0]=="✅ ALL OK"]
    elif filter_status == "❌ Ada Masalah":
        filtered_pos = [p for p in filtered_pos if "ISSUE" in str(sum_df[sum_df["PO"]==p]["Result"].values[0] if len(sum_df[sum_df["PO"]==p])>0 else "")]
    elif filter_status == "⚠️ Data Hilang":
        filtered_pos = [p for p in filtered_pos if "NO" in str(sum_df[sum_df["PO"]==p]["Result"].values[0] if len(sum_df[sum_df["PO"]==p])>0 else "")]

    st.markdown(f"<div class='section-label'>Menampilkan {len(filtered_pos)} dari {len(all_pos)} PO</div>", unsafe_allow_html=True)

    for po in filtered_pos:
        in_pdf = po in pdf_data
        in_xl  = po in xl_pos_set
        sd     = compare_po_size_detail(xl_df, pdf_data, po)
        info   = po_info(xl_df, pdf_data, po)

        n_match    = (sd["Status"]=="MATCH").sum()         if not sd.empty else 0
        n_mismatch = (sd["Status"]=="MISMATCH").sum()      if not sd.empty else 0
        n_only     = sd["Status"].str.contains("ONLY").sum() if not sd.empty else 0

        header_cls, header_html = render_po_header(po, n_match, n_mismatch, n_only, in_pdf, in_xl)

        expanded = (n_mismatch > 0 or n_only > 0) or (not in_pdf) or (not in_xl)

        with st.expander(f"PO {po}", expanded=expanded):
            st.markdown(f'<div class="po-card">{header_html}', unsafe_allow_html=True)

            if not in_pdf:
                st.warning("⚠️ PO ini tidak ditemukan di SAP PDF. Tidak ada data untuk dibandingkan.")
            elif not in_xl:
                st.warning("⚠️ PO ini tidak ditemukan di Infor Excel.")
            else:
                # Info strip
                st.markdown(render_info_strip(info), unsafe_allow_html=True)
                # Size table
                if not sd.empty:
                    st.markdown(render_size_table(sd, info), unsafe_allow_html=True)
                    # Infor total vs SAP total
                    xl_total  = int(xl_sizes_for_po(xl_df, po)["Infor Qty"].sum()) if not xl_sizes_for_po(xl_df, po).empty else 0
                    sap_total = int(pdf_data.get(po,{}).get("header",{}).get("Total Pairs",0) or 0)
                    diff_total = xl_total - sap_total
                    diff_color = "#10B981" if diff_total==0 else "#EF4444"
                    st.markdown(f"""
                    <div style="background:#F8FAFC;border-top:2px solid #E2E8F0;
                                padding:10px 16px;display:flex;gap:32px;
                                border-radius:0 0 12px 12px;font-size:0.82rem;">
                        <div>
                            <span style="color:#64748B;font-size:0.72rem;font-weight:600;text-transform:uppercase;">Total Infor</span><br>
                            <span style="font-size:1.1rem;font-weight:700;color:#1E293B;font-family:'IBM Plex Mono',monospace;">{xl_total:,} pairs</span>
                        </div>
                        <div>
                            <span style="color:#64748B;font-size:0.72rem;font-weight:600;text-transform:uppercase;">Total SAP</span><br>
                            <span style="font-size:1.1rem;font-weight:700;color:#1E293B;font-family:'IBM Plex Mono',monospace;">{sap_total:,} pairs</span>
                        </div>
                        <div>
                            <span style="color:#64748B;font-size:0.72rem;font-weight:600;text-transform:uppercase;">Selisih</span><br>
                            <span style="font-size:1.1rem;font-weight:700;color:{diff_color};font-family:'IBM Plex Mono',monospace;">{diff_total:+,} pairs</span>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.info("Tidak ada data ukuran untuk PO ini.")
            st.markdown("</div>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Discrepancies Only
# ═══════════════════════════════════════════════════════════════════════════════
with tab2:
    disc_found = False
    for po in all_pos:
        sd = compare_po_size_detail(xl_df, pdf_data, po)
        if sd.empty: continue
        issues = sd[sd["Status"]!="MATCH"]
        if issues.empty: continue
        disc_found = True
        info = po_info(xl_df, pdf_data, po)
        n_mm = (issues["Status"]=="MISMATCH").sum()
        n_only = issues["Status"].str.contains("ONLY").sum()
        cls = "bad" if n_mm>0 else "warn"
        st.markdown(f"""
        <div class="po-card" style="margin-bottom:16px;">
            <div class="po-card-header {cls}">
                <div class="po-number">PO {po}</div>
                <div>
                    {'<span class="badge-mismatch">'+str(n_mm)+' MISMATCH</span>' if n_mm>0 else ''}
                    {'<span class="badge-only-i" style="margin-left:6px;">'+str(n_only)+' ONLY-IN-ONE</span>' if n_only>0 else ''}
                </div>
            </div>
            {render_size_table(issues, info)}
        </div>
        """, unsafe_allow_html=True)

    if not disc_found:
        st.success("✅ Tidak ada discrepancy ditemukan — semua PO yang matched 100% OK!")

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 — Download
# ═══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown("### ⬇️ Download Excel Report")
    st.markdown("""
    Report Excel berisi 5 sheet:
    - **PO Compare Summary** — ringkasan per PO
    - **PO Compare Detail** — field-by-field per PO
    - **Size Detail** — semua kolom per ukuran
    - **Discrepancies Only** — hanya baris mismatch
    - **Raw Infor Data** — data mentah dari file Excel
    """)
    with st.spinner("Membuat Excel report..."):
        report_buf = build_report(xl_df, pdf_data, all_pos, xl_file.name, pdf_file.name)
    fname = f"InforVsSAP_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button(
        "📥 Download Excel Report",
        data=report_buf, file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, type="primary",
    )

    pdf_only = [p for p in pdf_pos if p not in xl_pos_set]
    xl_only  = [p for p in xl_pos  if p not in set(pdf_pos)]
    if pdf_only:
        st.warning(f"⚠️ {len(pdf_only)} PO ada di SAP tapi tidak di Infor: {', '.join(pdf_only)}")
    if xl_only:
        with st.expander(f"ℹ️ {len(xl_only)} PO ada di Infor tapi tidak di SAP PDF"):
            st.write(xl_only)
