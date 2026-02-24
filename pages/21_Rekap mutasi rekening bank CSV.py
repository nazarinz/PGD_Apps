import streamlit as st
import pandas as pd
import io
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Rekap Mutasi Rekening",
    page_icon="💳",
    layout="wide",
)

# ─────────────────────────────────────────────
# STYLING
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Plus Jakarta Sans', sans-serif;
}

/* Background */
.stApp {
    background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #0f172a 100%);
    min-height: 100vh;
}

/* Remove default padding */
.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
    max-width: 1100px;
}

/* Hero header */
.hero {
    text-align: center;
    padding: 3rem 2rem 2rem;
    margin-bottom: 2rem;
}
.hero-badge {
    display: inline-block;
    background: linear-gradient(90deg, #3b82f6, #06b6d4);
    color: white;
    font-size: 0.7rem;
    font-weight: 700;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    padding: 0.3rem 1rem;
    border-radius: 999px;
    margin-bottom: 1.2rem;
}
.hero h1 {
    font-size: 2.8rem;
    font-weight: 800;
    color: #f1f5f9;
    margin: 0 0 0.6rem;
    line-height: 1.15;
    letter-spacing: -0.02em;
}
.hero h1 span {
    background: linear-gradient(90deg, #3b82f6, #06b6d4);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}
.hero p {
    color: #94a3b8;
    font-size: 1.05rem;
    font-weight: 400;
    margin: 0;
}

/* Upload card */
.upload-card {
    background: rgba(30, 41, 59, 0.8);
    border: 1px solid rgba(59, 130, 246, 0.25);
    border-radius: 16px;
    padding: 2rem;
    margin-bottom: 1.5rem;
    backdrop-filter: blur(10px);
}

/* Metric cards */
.metric-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 1rem;
    margin-bottom: 1.5rem;
}
.metric-card {
    background: rgba(30, 41, 59, 0.9);
    border-radius: 12px;
    padding: 1.2rem 1.4rem;
    border: 1px solid rgba(255,255,255,0.07);
    position: relative;
    overflow: hidden;
}
.metric-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
}
.metric-card.saldo-awal::before  { background: #475569; }
.metric-card.kredit::before      { background: linear-gradient(90deg,#22c55e,#16a34a); }
.metric-card.debet::before       { background: linear-gradient(90deg,#ef4444,#dc2626); }
.metric-card.saldo-akhir::before { background: linear-gradient(90deg,#3b82f6,#06b6d4); }
.metric-label {
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #64748b;
    margin-bottom: 0.4rem;
}
.metric-value {
    font-family: 'JetBrains Mono', monospace;
    font-size: 1.05rem;
    font-weight: 600;
    color: #f1f5f9;
    word-break: break-all;
}
.metric-value.kredit  { color: #4ade80; }
.metric-value.debet   { color: #f87171; }
.metric-value.saldo-akhir { color: #60a5fa; }
.metric-sub {
    font-size: 0.72rem;
    color: #475569;
    margin-top: 0.25rem;
}

/* Day section header */
.day-header {
    font-size: 0.75rem;
    font-weight: 700;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #3b82f6;
    border-bottom: 1px solid rgba(59,130,246,0.2);
    padding-bottom: 0.4rem;
    margin: 1.5rem 0 0.8rem;
}

/* Table container */
.table-wrapper {
    background: rgba(15, 23, 42, 0.6);
    border-radius: 12px;
    border: 1px solid rgba(255,255,255,0.06);
    overflow: hidden;
    margin-bottom: 1.5rem;
}

/* Download button override */
.stDownloadButton > button {
    background: linear-gradient(135deg, #2563eb, #0891b2) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-weight: 700 !important;
    font-size: 0.95rem !important;
    padding: 0.7rem 2rem !important;
    width: 100% !important;
    transition: opacity 0.2s !important;
    letter-spacing: 0.01em;
}
.stDownloadButton > button:hover {
    opacity: 0.88 !important;
}

/* File uploader */
[data-testid="stFileUploader"] {
    background: rgba(15, 23, 42, 0.5);
    border: 2px dashed rgba(59, 130, 246, 0.35);
    border-radius: 12px;
    padding: 1rem;
}

/* Dataframe */
[data-testid="stDataFrame"] {
    border-radius: 8px;
    overflow: hidden;
}

/* Spinner */
.stSpinner > div { color: #3b82f6; }

/* Alert */
[data-testid="stAlert"] {
    border-radius: 10px;
    border: none;
}

/* Section title */
.section-title {
    font-size: 1rem;
    font-weight: 700;
    color: #e2e8f0;
    margin-bottom: 0.8rem;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}

/* Footer */
.footer {
    text-align: center;
    color: #334155;
    font-size: 0.75rem;
    margin-top: 3rem;
    padding-bottom: 1rem;
}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# PARSE CSV
# ─────────────────────────────────────────────
def parse_name(keterangan: str) -> str:
    """Extract clean sender/recipient name from keterangan field."""
    k = keterangan.strip()
    # BI-FAST pattern
    m = re.search(r'(?:TRANSFER\s+DR\s+\d+\s+|TANGGAL\s*:\S+\s+TRANSFER\s+DR\s+\d+\s+)(.+)', k, re.I)
    if m:
        return m.group(1).strip()
    # ESPAY / GoPay — extract TRFDN or sender label
    m2 = re.search(r'TRFDN-([A-Z ]+?)(?:ESPAY|$)', k, re.I)
    if m2:
        return m2.group(1).strip()
    m3 = re.search(r'GoPay Bank Transfe\S+\s+(.+)', k, re.I)
    if m3:
        return m3.group(1).strip()
    m4 = re.search(r'AIRPAY INTERNATION', k, re.I)
    if m4:
        return 'AIRPAY INTERNATIONAL'
    # E-BANKING: last uppercase token after amount
    m5 = re.search(r'\d{1,3}(?:\.\d{2})?(?:[A-Za-z\s\-_/\.\']+?)\s{2,}([A-Z][A-Z\s\.\',]+?)(?:,|$)', k)
    if m5:
        return m5.group(1).strip()
    # Fallback: last non-digit word group
    parts = re.split(r'\s{2,}', k)
    if len(parts) >= 2:
        return parts[-1].strip()
    return k[:40]


def parse_date_code(kode: str) -> str:
    """Convert DDMM or DD/MM code to readable date string."""
    kode = kode.strip()
    m = re.search(r'(\d{2})/(\d{2})', kode)
    if m:
        return f"{m.group(1)}/{m.group(2)}/2026"
    m2 = re.search(r'(\d{4})', kode)
    if m2:
        code = m2.group(1)
        return f"{code[:2]}/{code[2:]}/2026"
    return kode


def get_jenis(keterangan: str) -> str:
    k = keterangan.upper()
    if 'BI-FAST' in k:
        return 'BI-FAST'
    if 'ESPAY' in k or 'DOMPET ANAK BANGSA' in k or 'AIRPAY' in k or 'GOPAY' in k:
        return 'E-Banking (FinTech)'
    if 'E-BANKING' in k:
        return 'E-Banking'
    return 'Transfer'


def parse_csv(content: str):
    lines = [l.strip() for l in content.splitlines() if l.strip()]

    # Detect footer summary lines
    summary = {}
    transactions = []

    for line in lines:
        low = line.lower()
        if low.startswith('saldo awal'):
            parts = line.split(',')
            summary['saldo_awal'] = float(parts[-1].replace("'", '').strip())
        elif low.startswith('kredit'):
            parts = line.split(',')
            summary['kredit'] = float(parts[-1].replace("'", '').strip())
        elif low.startswith('debet'):
            parts = line.split(',')
            summary['debet'] = float(parts[-1].replace("'", '').strip())
        elif low.startswith('saldo akhir'):
            parts = line.split(',')
            summary['saldo_akhir'] = float(parts[-1].replace("'", '').strip())
        elif low.startswith("'pend") or low.startswith("pend"):
            # Parse transaction line: split carefully
            # Format: 'PEND,KETERANGAN,'CABANG,JUMLAH,,SALDO
            # Remove leading quote
            clean = line.lstrip("'")
            parts = clean.split(',')
            if len(parts) < 5:
                continue
            keterangan = parts[1]
            # Find numeric fields from the end
            try:
                saldo = float(parts[-1])
                tipe = parts[-2].strip()  # CR or DB
                jumlah = float(parts[-3])
                # date from keterangan
                date_str = parse_date_code(keterangan)
                nama = parse_name(keterangan)
                jenis = get_jenis(keterangan)
                transactions.append({
                    'Tanggal': date_str,
                    'Jenis': jenis,
                    'Pengirim/Penerima': nama,
                    'Tipe': tipe,
                    'Jumlah': jumlah,
                    'Saldo': saldo,
                    'Keterangan_Raw': keterangan,
                })
            except (ValueError, IndexError):
                continue

    return transactions, summary


# ─────────────────────────────────────────────
# GENERATE EXCEL
# ─────────────────────────────────────────────
def generate_excel(transactions: list, summary: dict, filename: str) -> bytes:
    wb = Workbook()
    DARK   = "1F3864"
    MED    = "2E75B6"
    LGREY  = "F7F9FC"
    GREEN  = "E2EFDA"
    RED    = "FCE4D6"
    LBLUE  = "DBEAFE"
    WHITE  = "FFFFFF"
    thin   = Side(style='thin', color='D1D5DB')
    bdr    = Border(left=thin, right=thin, top=thin, bottom=thin)

    hf  = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
    df  = Font(name='Calibri', size=9)
    bf  = Font(name='Calibri', bold=True, size=9)
    tf  = Font(name='Calibri', bold=True, size=13, color='FFFFFF')

    def hdr_cell(ws, row, col, val, width_hint=None):
        c = ws.cell(row=row, column=col, value=val)
        c.font = hf
        c.fill = PatternFill('solid', start_color=MED)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = bdr
        return c

    def title_row(ws, row, ncols, text):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
        c = ws.cell(row=row, column=1, value=text)
        c.font = tf
        c.fill = PatternFill('solid', start_color=DARK)
        c.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[row].height = 28

    # ── Sheet 1: Detail ──
    ws1 = wb.active
    ws1.title = "Detail Mutasi"
    ws1.freeze_panes = 'A4'

    periode = filename.replace('.csv', '').replace('_', ' ')
    title_row(ws1, 1, 7, f'REKAP MUTASI REKENING')
    ws1.merge_cells('A2:G2')
    c2 = ws1.cell(row=2, column=1, value=f'File: {filename}   |   Dibuat: {datetime.now().strftime("%d %b %Y %H:%M")}')
    c2.font = Font(name='Calibri', italic=True, size=9, color='FFFFFF')
    c2.fill = PatternFill('solid', start_color=MED)
    c2.alignment = Alignment(horizontal='center', vertical='center')
    ws1.row_dimensions[2].height = 16

    hdrs = ['No', 'Tanggal', 'Jenis', 'Pengirim / Penerima', 'Debet (Rp)', 'Kredit (Rp)', 'Saldo (Rp)']
    ws1.row_dimensions[3].height = 22
    for ci, h in enumerate(hdrs, 1):
        hdr_cell(ws1, 3, ci, h)

    for i, row in enumerate(transactions, 1):
        r = i + 3
        is_db = row['Tipe'] == 'DB'
        debet  = row['Jumlah'] if is_db else None
        kredit = row['Jumlah'] if not is_db else None
        bg = LGREY if i % 2 == 0 else WHITE

        vals = [i, row['Tanggal'], row['Jenis'], row['Pengirim/Penerima'], debet, kredit, row['Saldo']]
        for ci, val in enumerate(vals, 1):
            cell = ws1.cell(row=r, column=ci, value=val)
            cell.font = df
            cell.border = bdr
            if ci == 5 and val:
                cell.fill = PatternFill('solid', start_color=RED)
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal='right')
            elif ci == 6 and val:
                cell.fill = PatternFill('solid', start_color=GREEN)
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal='right')
            elif ci == 7:
                cell.fill = PatternFill('solid', start_color=bg)
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal='right')
            elif ci == 1:
                cell.fill = PatternFill('solid', start_color=bg)
                cell.alignment = Alignment(horizontal='center')
            else:
                cell.fill = PatternFill('solid', start_color=bg)

    ws1.column_dimensions['A'].width = 5
    ws1.column_dimensions['B'].width = 11
    ws1.column_dimensions['C'].width = 20
    ws1.column_dimensions['D'].width = 30
    ws1.column_dimensions['E'].width = 16
    ws1.column_dimensions['F'].width = 16
    ws1.column_dimensions['G'].width = 18

    # ── Sheet 2: Ringkasan ──
    ws2 = wb.create_sheet("Ringkasan")

    title_row(ws2, 1, 4, 'RINGKASAN MUTASI REKENING')
    ws2.merge_cells('A2:D2')
    c2b = ws2.cell(row=2, column=1, value=f'Dibuat otomatis dari file: {filename}')
    c2b.font = Font(name='Calibri', italic=True, size=9, color='FFFFFF')
    c2b.fill = PatternFill('solid', start_color=MED)
    c2b.alignment = Alignment(horizontal='center', vertical='center')
    ws2.row_dimensions[2].height = 16

    def kv(ws, row, label, value, bg=WHITE, bold=False, is_num=True):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        lc = ws.cell(row=row, column=1, value=label)
        lc.font = bf if bold else df
        lc.fill = PatternFill('solid', start_color=bg)
        lc.border = bdr
        vc = ws.cell(row=row, column=3, value=value)
        vc.font = bf if bold else df
        vc.fill = PatternFill('solid', start_color=bg)
        vc.border = bdr
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=4)
        if is_num:
            vc.number_format = '#,##0.00'
            vc.alignment = Alignment(horizontal='right')

    # Ikhtisar saldo
    ws2.merge_cells('A4:D4')
    sh = ws2.cell(row=4, column=1, value='IKHTISAR SALDO')
    sh.font = hf; sh.fill = PatternFill('solid', start_color=DARK)
    sh.alignment = Alignment(horizontal='center'); ws2.row_dimensions[4].height = 20

    kv(ws2, 5, 'Saldo Awal', summary.get('saldo_awal', 0), bg=LGREY)
    kv(ws2, 6, 'Total Kredit (Masuk)', summary.get('kredit', 0), bg=GREEN)
    kv(ws2, 7, 'Total Debet (Keluar)', summary.get('debet', 0), bg=RED)
    kv(ws2, 8, 'Saldo Akhir', summary.get('saldo_akhir', 0), bg=LBLUE, bold=True)

    # Per-hari
    dates = sorted(set(r['Tanggal'] for r in transactions))
    ws2.merge_cells('A10:D10')
    sh2 = ws2.cell(row=10, column=1, value='RINGKASAN PER HARI')
    sh2.font = hf; sh2.fill = PatternFill('solid', start_color=DARK)
    sh2.alignment = Alignment(horizontal='center'); ws2.row_dimensions[10].height = 20

    day_hdrs = ['Tanggal', 'Tx Masuk', 'Total Masuk (Rp)', 'Total Keluar (Rp)']
    for ci, h in enumerate(day_hdrs, 1):
        hdr_cell(ws2, 11, ci, h)

    for ri, d in enumerate(dates):
        r = 12 + ri
        cr_amt = sum(t['Jumlah'] for t in transactions if t['Tanggal']==d and t['Tipe']=='CR')
        db_amt = sum(t['Jumlah'] for t in transactions if t['Tanggal']==d and t['Tipe']=='DB')
        cr_cnt = sum(1 for t in transactions if t['Tanggal']==d and t['Tipe']=='CR')
        bg = LGREY if ri % 2 == 0 else WHITE
        for ci, val in enumerate([d, cr_cnt, cr_amt, db_amt], 1):
            c = ws2.cell(row=r, column=ci, value=val)
            c.font = df; c.border = bdr
            c.fill = PatternFill('solid', start_color=bg)
            if ci in (3, 4):
                c.number_format = '#,##0'; c.alignment = Alignment(horizontal='right')
            elif ci == 2:
                c.alignment = Alignment(horizontal='center')

    # Top 10
    top_cr = sorted([(t['Jumlah'], t['Pengirim/Penerima'], t['Tanggal']) for t in transactions if t['Tipe']=='CR'], reverse=True)[:10]
    start_r = 12 + len(dates) + 2
    ws2.merge_cells(f'A{start_r}:D{start_r}')
    sh3 = ws2.cell(row=start_r, column=1, value='TOP 10 TRANSAKSI MASUK TERBESAR')
    sh3.font = hf; sh3.fill = PatternFill('solid', start_color=DARK)
    sh3.alignment = Alignment(horizontal='center'); ws2.row_dimensions[start_r].height = 20

    for ci, h in enumerate(['No','Pengirim','Tanggal','Jumlah (Rp)'], 1):
        hdr_cell(ws2, start_r+1, ci, h)

    for ri, (amt, nm, tgl) in enumerate(top_cr, 1):
        r = start_r + 1 + ri
        bg = GREEN if ri % 2 == 0 else WHITE
        for ci, val in enumerate([ri, nm, tgl, amt], 1):
            c = ws2.cell(row=r, column=ci, value=val)
            c.font = df; c.border = bdr
            c.fill = PatternFill('solid', start_color=bg if ci != 4 else GREEN)
            if ci == 4:
                c.number_format = '#,##0'; c.alignment = Alignment(horizontal='right')
            elif ci == 1:
                c.alignment = Alignment(horizontal='center')

    ws2.column_dimensions['A'].width = 20
    ws2.column_dimensions['B'].width = 15
    ws2.column_dimensions['C'].width = 20
    ws2.column_dimensions['D'].width = 20

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────
# FORMAT HELPERS
# ─────────────────────────────────────────────
def fmt_rp(val: float) -> str:
    return f"Rp {val:,.0f}".replace(",", ".")


# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <div class="hero-badge">💳 Bank Mutation Tools</div>
    <h1>Rekap Mutasi <span>Rekening</span></h1>
    <p>Upload file CSV mutasi rekening → otomatis jadi Excel terstruktur & rapi</p>
</div>
""", unsafe_allow_html=True)

# Upload zone
with st.container():
    st.markdown('<div class="upload-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📂 Upload File CSV Mutasi</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "Drag & drop atau klik untuk upload",
        type=['csv'],
        label_visibility='collapsed'
    )
    st.markdown('<p style="color:#475569;font-size:0.8rem;margin-top:0.5rem;">Format: Mutasi rekening BCA / bank umum. Bisa upload lebih dari satu file sekaligus untuk penggabungan.</p>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

if uploaded:
    raw = uploaded.read().decode('utf-8', errors='replace')
    fname = uploaded.name

    with st.spinner("Memproses data mutasi..."):
        transactions, summary = parse_csv(raw)

    if not transactions:
        st.error("⚠️ Tidak ada transaksi yang berhasil diparsing. Pastikan format CSV sesuai.")
        st.stop()

    cr_total = summary.get('kredit', sum(t['Jumlah'] for t in transactions if t['Tipe']=='CR'))
    db_total = summary.get('debet',  sum(t['Jumlah'] for t in transactions if t['Tipe']=='DB'))
    sa = summary.get('saldo_awal', 0)
    se = summary.get('saldo_akhir', 0)
    cr_cnt = sum(1 for t in transactions if t['Tipe']=='CR')
    db_cnt = sum(1 for t in transactions if t['Tipe']=='DB')

    # ── Metrics ──
    st.markdown(f"""
    <div class="metric-grid">
        <div class="metric-card saldo-awal">
            <div class="metric-label">Saldo Awal</div>
            <div class="metric-value">{fmt_rp(sa)}</div>
        </div>
        <div class="metric-card kredit">
            <div class="metric-label">Total Masuk</div>
            <div class="metric-value kredit">{fmt_rp(cr_total)}</div>
            <div class="metric-sub">{cr_cnt} transaksi</div>
        </div>
        <div class="metric-card debet">
            <div class="metric-label">Total Keluar</div>
            <div class="metric-value debet">{fmt_rp(db_total)}</div>
            <div class="metric-sub">{db_cnt} transaksi</div>
        </div>
        <div class="metric-card saldo-akhir">
            <div class="metric-label">Saldo Akhir</div>
            <div class="metric-value saldo-akhir">{fmt_rp(se)}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Preview table per tanggal ──
    df = pd.DataFrame(transactions)
    dates = sorted(df['Tanggal'].unique())

    for d in dates:
        df_day = df[df['Tanggal'] == d].copy()
        cr_d = df_day[df_day['Tipe']=='CR']['Jumlah'].sum()
        db_d = df_day[df_day['Tipe']=='DB']['Jumlah'].sum()

        st.markdown(f"""
        <div class="day-header">
            📅 {d} &nbsp;&nbsp;·&nbsp;&nbsp; 
            Masuk: {fmt_rp(cr_d)} &nbsp;·&nbsp; Keluar: {fmt_rp(db_d)}
        </div>
        """, unsafe_allow_html=True)

        preview_df = df_day[['Tipe','Pengirim/Penerima','Jenis','Jumlah','Saldo']].copy()
        preview_df['Tipe'] = preview_df['Tipe'].map({'CR': '✅ CR', 'DB': '🔴 DB'})
        preview_df.columns = ['Tipe','Pengirim/Penerima','Jenis','Jumlah (Rp)','Saldo (Rp)']
        preview_df = preview_df.reset_index(drop=True)
        preview_df.index = preview_df.index + 1

        st.dataframe(
            preview_df.style.format({'Jumlah (Rp)': '{:,.0f}', 'Saldo (Rp)': '{:,.0f}'}),
            use_container_width=True,
            height=min(400, 38 * (len(preview_df) + 1))
        )

    # ── Download ──
    st.markdown("---")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        out_name = fname.replace('.csv', '_REKAP.xlsx')
        excel_bytes = generate_excel(transactions, summary, fname)
        st.download_button(
            label=f"⬇️  Download Excel — {out_name}",
            data=excel_bytes,
            file_name=out_name,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )

else:
    # Placeholder state
    st.markdown("""
    <div style="text-align:center;padding:3rem 1rem;color:#334155;">
        <div style="font-size:3rem;margin-bottom:1rem;">☁️</div>
        <div style="font-size:0.95rem;font-weight:500;">Belum ada file yang diupload</div>
        <div style="font-size:0.8rem;margin-top:0.4rem;color:#475569;">Upload file CSV mutasi rekening untuk memulai</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown('<div class="footer">Rekap Mutasi Rekening · Dibuat dengan Streamlit & openpyxl</div>', unsafe_allow_html=True)
