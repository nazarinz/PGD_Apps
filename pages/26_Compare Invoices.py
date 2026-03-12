"""
Invoice Comparison Tool — FMS vs FVB
Streamlit App

Run:
    streamlit run app.py
"""

import io
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── CONFIG ────────────────────────────────────────────────────────────────────

log = logging.getLogger(__name__)

HEADER_FIELDS = [
    "Invoice No.", "Export Type", "Brand Name", "Buyer Name",
    "Consignee", "Ex-Go Date", "PEB No", "PEB Date", "Amount", "Gross Weight"
]
COMPARE_HEADER_FIELDS = ["Amount", "Gross Weight"]

ITEM_COLS = [
    "Item Seq.", "Fact No", "SO / PO No", "Mat No", "Mat Name",
    "Hscode No", "Qty", "Jenis Satuan", "Unit Price", "Amount",
    "FOC Mark", "Pack Qty", "Pack Name", "Net Weight", "Gross Weight", "Volume"
]
SKIP_ITEM_COMPARE = {"SO / PO No"}
COMPARE_ITEM_COLS = [c for c in ITEM_COLS if c not in SKIP_ITEM_COMPARE]

# Colors
C_FMS_HEAD  = "1E3A5F"
C_FVB_HEAD  = "7C4A00"
C_FMS_LIGHT = "D6E4F0"
C_FVB_LIGHT = "FFF3CD"
C_GREEN     = "D4EDDA"
C_RED       = "F8D7DA"
C_YELLOW    = "FFF9C4"
C_WHITE     = "FFFFFF"
C_FMS_COL   = "2471A3"
C_FVB_COL   = "D68910"
C_SUMMARY   = "1A1A2E"

# ── DATA MODELS ───────────────────────────────────────────────────────────────

@dataclass
class InvoiceData:
    invoice_no: str
    header: dict[str, str] = field(default_factory=dict)
    items: list[dict[str, Any]] = field(default_factory=list)
    total_row: list[Any] | None = None
    source_file: str = ""

@dataclass
class Diff:
    invoice_no: str
    scope: str
    field: str
    item_seq: str
    fms_val: str
    fvb_val: str

# ── PARSER ────────────────────────────────────────────────────────────────────

def _norm(val: Any) -> str:
    if val is None:
        return ""
    if isinstance(val, float) and not (val != val):  # not NaN
        if val == int(val):
            return str(int(val))
    return str(val).strip()


def parse_invoice(file_obj, filename: str) -> InvoiceData:
    df_raw = pd.read_excel(file_obj, header=None, dtype=str)
    rows = df_raw.fillna("").values.tolist()

    header: dict[str, str] = {}
    item_header_row = -1

    for i, row in enumerate(rows):
        for j, cell in enumerate(row):
            cell_s = str(cell).strip().rstrip(":")
            for f in HEADER_FIELDS:
                if cell_s == f or cell_s == f.rstrip("."):
                    val = _norm(rows[i][j + 1] if j + 1 < len(rows[i]) else "")
                    header[f] = val
        if any(str(c).strip() == "Item Seq." for c in row):
            item_header_row = i

    items: list[dict[str, Any]] = []
    total_row = None

    if item_header_row >= 0:
        col_map = {str(rows[item_header_row][k]).strip(): k
                   for k in range(len(rows[item_header_row]))}
        for row in rows[item_header_row + 1:]:
            if all(str(c).strip() == "" for c in row):
                break
            seq = str(row[0]).strip() if row else ""
            if seq == "" and any(str(c).strip() != "" for c in row):
                total_row = row
                continue
            if seq and not seq.replace(".", "").isdigit():
                break
            item = {col: _norm(row[col_map[col]]) if col in col_map else ""
                    for col in ITEM_COLS}
            items.append(item)

    inv_no = header.get("Invoice No.", Path(filename).stem)
    return InvoiceData(invoice_no=inv_no, header=header, items=items,
                       total_row=total_row, source_file=filename)


def compare_pair(fms: InvoiceData, fvb: InvoiceData) -> list[Diff]:
    diffs: list[Diff] = []
    inv = fms.invoice_no

    for f in COMPARE_HEADER_FIELDS:
        fv, bv = fms.header.get(f, ""), fvb.header.get(f, "")
        if fv != bv:
            diffs.append(Diff(inv, "header", f, "", fv, bv))

    for idx, fms_item in enumerate(fms.items):
        fvb_item = fvb.items[idx] if idx < len(fvb.items) else {}
        seq = fms_item.get("Item Seq.", str(idx + 1))
        for col in COMPARE_ITEM_COLS:
            fv = fms_item.get(col, "")
            bv = _norm(fvb_item.get(col, ""))
            if fv != bv:
                diffs.append(Diff(inv, "item", col, seq, fv, bv))

    if fms.total_row and fvb.total_row:
        for ci, (fv, bv) in enumerate(zip(fms.total_row, fvb.total_row)):
            if _norm(fv) != _norm(bv) and (_norm(fv) or _norm(bv)):
                col_name = ITEM_COLS[ci] if ci < len(ITEM_COLS) else f"Col{ci}"
                diffs.append(Diff(inv, "total", col_name, "TOTAL", _norm(fv), _norm(bv)))

    return diffs

# ── EXCEL BUILDER ─────────────────────────────────────────────────────────────

def _fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)

def _border() -> Border:
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def _font(bold=True, size=10, color="000000") -> Font:
    return Font(bold=bold, size=size, color=color)


def _write_header_block(ws, r: int, label: str, data: InvoiceData,
                         lc: str, fc: str) -> int:
    ws.cell(r, 1, label).font = _font(color="FFFFFF", size=11)
    ws.cell(r, 1).fill = _fill(lc)
    ws.cell(r, 1).alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    for f in HEADER_FIELDS:
        r += 1
        ws.cell(r, 2, f"{f} :").font = _font()
        ws.cell(r, 2).fill = _fill(fc)
        ws.cell(r, 3, data.header.get(f, "")).fill = _fill(C_WHITE)
    return r + 1


def _write_items_block(ws, r: int, data: InvoiceData,
                        col_color: str, diffs: list[Diff]) -> int:
    diff_keys = {(d.item_seq, d.field) for d in diffs if d.scope == "item"}
    diff_total = {d.field for d in diffs if d.scope == "total"}

    for ci, col in enumerate(ITEM_COLS, 1):
        c = ws.cell(r, ci, col)
        c.font = _font(color="FFFFFF", size=9)
        c.fill = _fill(col_color)
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        c.border = _border()
    r += 1

    for item in data.items:
        for ci, col in enumerate(ITEM_COLS, 1):
            c = ws.cell(r, ci, item.get(col, ""))
            c.border = _border()
            c.alignment = Alignment(horizontal="center")
            key = (item.get("Item Seq.", ""), col)
            if key in diff_keys:
                c.fill = _fill(C_RED)
            elif col in SKIP_ITEM_COMPARE:
                c.fill = _fill(C_YELLOW)
        r += 1

    if data.total_row:
        for ci, val in enumerate(data.total_row[:len(ITEM_COLS)], 1):
            col_name = ITEM_COLS[ci - 1] if ci - 1 < len(ITEM_COLS) else ""
            c = ws.cell(r, ci, _norm(val))
            c.font = _font(size=9)
            c.border = _border()
            c.alignment = Alignment(horizontal="center")
            if col_name in diff_total:
                c.fill = _fill(C_RED)
        r += 1

    return r + 1


def _write_ok_row(ws, r: int, diffs: list[Diff]) -> None:
    diff_cols = {d.field for d in diffs if d.scope in ("item", "total")}
    ws.cell(r, 1, "CHECK").font = _font(color="FFFFFF", size=9)
    ws.cell(r, 1).fill = _fill("444444")
    ws.cell(r, 1).alignment = Alignment(horizontal="center")
    for ci, col in enumerate(ITEM_COLS, 1):
        if col in SKIP_ITEM_COMPARE:
            lbl, clr = "SKIP", C_YELLOW
        elif col in diff_cols:
            lbl, clr = "DIFF", C_RED
        else:
            lbl, clr = "OK", C_GREEN
        c = ws.cell(r, ci, lbl)
        c.font = _font(size=9)
        c.fill = _fill(clr)
        c.alignment = Alignment(horizontal="center")
        c.border = _border()


def build_excel(pairs: list[tuple[InvoiceData, InvoiceData, list[Diff]]]) -> bytes:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    summary_rows: list[tuple] = []
    col_widths = [8,18,36,16,30,12,8,10,10,12,8,8,8,11,12,9]

    for fms, fvb, diffs in pairs:
        safe = fms.invoice_no.replace("/","_").replace("\\","_")[:31]
        ws = wb.create_sheet(safe)
        for ci, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(ci)].width = w

        r = 1
        r = _write_header_block(ws, r, "FMS", fms, C_FMS_COL, C_FMS_LIGHT)
        r = _write_items_block(ws, r, fms, C_FMS_COL, diffs)
        r = _write_header_block(ws, r, "FVB", fvb, C_FVB_COL, C_FVB_LIGHT)
        r = _write_items_block(ws, r, fvb, C_FVB_COL, diffs)
        _write_ok_row(ws, r, diffs)

        for d in diffs:
            summary_rows.append((d.invoice_no, d.scope, d.field,
                                  d.item_seq, d.fms_val, d.fvb_val))

    # Summary sheet
    ws_s = wb.create_sheet("SUMMARY")
    hdrs = ["Invoice No.", "Scope", "Field", "Item Seq.", "FMS Value", "FVB Value"]
    for ci, h in enumerate(hdrs, 1):
        c = ws_s.cell(1, ci, h)
        c.font = _font(color="FFFFFF")
        c.fill = _fill(C_SUMMARY)
        c.alignment = Alignment(horizontal="center")
    ws_s.column_dimensions["A"].width = 22
    ws_s.column_dimensions["B"].width = 10
    ws_s.column_dimensions["C"].width = 20
    ws_s.column_dimensions["D"].width = 10
    ws_s.column_dimensions["E"].width = 28
    ws_s.column_dimensions["F"].width = 28

    if not summary_rows:
        ws_s.cell(2, 1, "✓ No differences found").font = Font(color="2E7D32", bold=True)
    else:
        for ri, row in enumerate(summary_rows, 2):
            for ci, val in enumerate(row, 1):
                c = ws_s.cell(ri, ci, val)
                c.border = _border()
                c.fill = _fill(C_RED)
                c.alignment = Alignment(horizontal="center")
    ws_s.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ── STREAMLIT UI ──────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Invoice Compare — FMS vs FVB",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #0f1117; }
[data-testid="stHeader"] { background: transparent; }
.main-title { font-size: 2rem; font-weight: 800; color: #fff;
              font-family: 'JetBrains Mono', monospace; margin-bottom: 0; }
.sub-title  { font-size: 0.85rem; color: #666; margin-bottom: 2rem;
              font-family: monospace; letter-spacing: 2px; }
.metric-box { background: #1a1a2e; border: 1px solid #333; border-radius: 10px;
              padding: 18px 22px; text-align: center; }
.metric-val { font-size: 2.2rem; font-weight: 800; font-family: monospace; }
.metric-lbl { font-size: 0.7rem; color: #666; letter-spacing: 3px; margin-top: 4px; }
.inv-card   { background: #1a1a2e; border: 1px solid #2a2a4a; border-radius: 10px;
              padding: 14px 18px; margin-bottom: 10px; }
.inv-ok     { border-left: 4px solid #4ade80; }
.inv-diff   { border-left: 4px solid #f87171; }
.badge-ok   { background: #14532d; color: #86efac; border-radius: 4px;
              padding: 2px 10px; font-size: 0.72rem; font-weight: 700; font-family: monospace; }
.badge-diff { background: #7f1d1d; color: #fca5a5; border-radius: 4px;
              padding: 2px 10px; font-size: 0.72rem; font-weight: 700; font-family: monospace; }
.diff-row   { display: grid; grid-template-columns: 170px 1fr 1fr; gap: 10px;
              font-size: 0.8rem; margin-top: 6px; }
.diff-field { color: #f59e0b; font-weight: 700; font-family: monospace; }
.diff-fms   { color: #60a5fa; font-family: monospace; word-break: break-all; }
.diff-fvb   { color: #fb923c; font-family: monospace; word-break: break-all; }
.section-lbl{ font-size: 0.65rem; letter-spacing: 4px; color: #555;
              text-transform: uppercase; margin: 1.5rem 0 0.5rem; font-family: monospace; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 Invoice Compare</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">FMS vs FVB — RECONCILIATION TOOL</div>', unsafe_allow_html=True)

# ── UPLOAD ────────────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.markdown("#### 🔵 FMS Files")
    fms_uploads = st.file_uploader("Upload FMS Excel", type=["xlsx", "xls"],
                                    accept_multiple_files=True, key="fms",
                                    label_visibility="collapsed")
    if fms_uploads:
        for f in fms_uploads:
            st.caption(f"✓ {f.name}")

with col2:
    st.markdown("#### 🟡 FVB Files")
    fvb_uploads = st.file_uploader("Upload FVB Excel", type=["xlsx", "xls"],
                                    accept_multiple_files=True, key="fvb",
                                    label_visibility="collapsed")
    if fvb_uploads:
        for f in fvb_uploads:
            st.caption(f"✓ {f.name}")

st.divider()

# ── PROCESS ───────────────────────────────────────────────────────────────────
if not fms_uploads or not fvb_uploads:
    st.info("⬆ Upload setidaknya 1 file FMS dan 1 file FVB untuk memulai.")
    st.stop()

run_btn = st.button("▶ COMPARE", type="primary", use_container_width=False)

if run_btn or st.session_state.get("result"):

    if run_btn:
        with st.spinner("Memproses file..."):
            # Parse
            fms_map: dict[str, InvoiceData] = {}
            fvb_map: dict[str, InvoiceData] = {}
            parse_errors: list[str] = []

            for f in fms_uploads:
                try:
                    d = parse_invoice(f, f.name)
                    fms_map[d.invoice_no] = d
                except Exception as e:
                    parse_errors.append(f"FMS {f.name}: {e}")

            for f in fvb_uploads:
                try:
                    d = parse_invoice(f, f.name)
                    fvb_map[d.invoice_no] = d
                except Exception as e:
                    parse_errors.append(f"FVB {f.name}: {e}")

            # Match & compare
            all_keys = sorted(set(fms_map) | set(fvb_map))
            pairs: list[tuple] = []
            missing: list[dict] = []

            for key in all_keys:
                if key not in fms_map:
                    missing.append({"invoice": key, "reason": "FMS tidak ditemukan"})
                elif key not in fvb_map:
                    missing.append({"invoice": key, "reason": "FVB tidak ditemukan"})
                else:
                    diffs = compare_pair(fms_map[key], fvb_map[key])
                    pairs.append((fms_map[key], fvb_map[key], diffs))

            st.session_state["result"] = {
                "pairs": pairs, "missing": missing,
                "parse_errors": parse_errors
            }

    res = st.session_state.get("result", {})
    pairs     = res.get("pairs", [])
    missing   = res.get("missing", [])
    p_errors  = res.get("parse_errors", [])

    # ── ERRORS ────────────────────────────────────────────────────────────────
    for err in p_errors:
        st.error(f"⚠ {err}")

    if not pairs and not missing:
        st.warning("Tidak ada invoice yang cocok antara FMS dan FVB.")
        st.stop()

    # ── METRICS ───────────────────────────────────────────────────────────────
    total_diffs = sum(len(d) for _, _, d in pairs)
    ok_count    = sum(1 for _, _, d in pairs if not d)
    diff_count  = len(pairs) - ok_count

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Invoice Cocok",   len(pairs))
    m2.metric("✅ Semua OK",     ok_count)
    m3.metric("❌ Ada Beda",     diff_count)
    m4.metric("Total Perbedaan", total_diffs)

    st.divider()

    # ── DOWNLOAD BUTTON ───────────────────────────────────────────────────────
    if pairs:
        with st.spinner("Membuat file Excel..."):
            excel_bytes = build_excel(pairs)
        st.download_button(
            label="⬇ Download Hasil Compare (.xlsx)",
            data=excel_bytes,
            file_name="Invoice_Compare_Result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    st.divider()

    # ── INVOICE LIST ──────────────────────────────────────────────────────────
    st.markdown('<div class="section-lbl">Hasil Per Invoice</div>', unsafe_allow_html=True)

    for fms, fvb, diffs in pairs:
        card_cls = "inv-diff" if diffs else "inv-ok"
        badge    = f'<span class="badge-diff">{len(diffs)} DIFF</span>' if diffs \
                   else '<span class="badge-ok">✓ OK</span>'

        with st.expander(f"{'🔴' if diffs else '🟢'}  {fms.invoice_no}  —  "
                         f"FMS: {len(fms.items)} item | FVB: {len(fvb.items)} item"
                         + (f"  •  {len(diffs)} perbedaan" if diffs else "  •  semua sama")):

            if not diffs:
                st.success("Tidak ada perbedaan pada invoice ini.")
            else:
                # Group by scope
                header_diffs = [d for d in diffs if d.scope == "header"]
                item_diffs   = [d for d in diffs if d.scope in ("item", "total")]

                if header_diffs:
                    st.markdown("**Header**")
                    hdf = pd.DataFrame([{
                        "Field": d.field,
                        "FMS": d.fms_val,
                        "FVB": d.fvb_val
                    } for d in header_diffs])
                    st.dataframe(hdf, use_container_width=True, hide_index=True)

                if item_diffs:
                    st.markdown("**Items**")
                    idf = pd.DataFrame([{
                        "Item Seq.": d.item_seq,
                        "Field": d.field,
                        "FMS": d.fms_val,
                        "FVB": d.fvb_val
                    } for d in item_diffs])
                    st.dataframe(idf, use_container_width=True, hide_index=True)

    # ── MISSING ───────────────────────────────────────────────────────────────
    if missing:
        st.divider()
        st.markdown('<div class="section-lbl">Invoice Tidak Cocok</div>', unsafe_allow_html=True)
        for m in missing:
            st.warning(f"⚠  **{m['invoice']}** — {m['reason']}")
