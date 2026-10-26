"""
HRIS Monthly Attendance ETL (Streamlit)
---------------------------------------
Upload PDF "Monthly Attendance Report" HRIS (PT.GLOSTAR INDONESIA) -> parse -> QA -> Excel.

Run:
    pip install streamlit pdfplumber pandas openpyxl
    streamlit run hris_attendance_etl_app.py

Output workbook: README, Attendance_Daily, Employee_Monthly, Dim_Employee, Dim_Status, QA_Checks
"""
import re
import io
import calendar
import datetime as dt
from collections import Counter

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# Config
# ----------------------------------------------------------------------------
# Kolom ringkasan per karyawan di PDF (urut dari kiri ke kanan). Januari punya 18 kolom, Februari 19 (+SOR).
SUM_COLS = ['H', 'A', 'I', 'CT', 'S', 'SD', 'TL', 'IR', 'DL', 'CH', 'CHK', 'HD', 'PC', 'LC',
            'N', 'CL4', 'CL5', 'CL6', 'SOR']

STATUS = {  # kode: (deskripsi, grup)
    'N': ('Normal', 'Hadir'), 'TL': ('Terlambat', 'Hadir'), 'PC': ('Pulang Cepat', 'Hadir'),
    'DL': ('Dinas Luar', 'Hadir'),
    'CT': ('Cuti', 'Cuti'), 'CH': ('Cuti Hamil', 'Cuti'),
    'SD': ('Surat Dokter', 'Sakit'), 'S': ('Sakit', 'Sakit'),
    'I': ('Izin', 'Izin'), 'IR': ('Izin Resmi', 'Izin'),
    'A': ('Alpa', 'Alpa'), 'H': ('Hari Libur', 'Libur'),
    'OFF': ('Weekend (tidak ada kode di PDF)', 'Weekend'),
    'NODATA': ('Hari kerja tanpa kode di PDF', 'No Data'),
}
# Koreksi manual untuk hari kerja yang kosong di PDF HRIS (sudah dikonfirmasi user).
# Format: (NIK, 'YYYY-MM-DD'): status. Hanya berlaku jika sel di PDF memang kosong.
MANUAL_OVERRIDES = {
    ('AD25722', '2025-06-05'): 'DL',   # Ahmad Ilham Fauzi, hari ke-4 dinas luar (Juni 2025)
    ('AD32820', '2025-05-22'): 'DL',   # Farhan Nandyawan Syarief, hari ke-4 dinas luar (Mei 2025)
}
RAW_ALIAS = {'I*': 'IR'}  # di ringkasan HRIS, I* dihitung sebagai IR

# Batas koordinat (pt) untuk layout PDF ini
X_DAY_MIN, X_DAY_MAX = 150, 570       # area kolom tanggal
X_SUMMARY_MIN = 570                    # area kolom ringkasan
X_TOTAL_SPLIT = 568.5                  # pemisah kolom hari-31 dan kolom total lembur
X_NAME_MIN, X_NAME_MAX = 66, 300       # area nama
X_NAME_JUNK_SPACE = 82                 # spasi < x ini adalah sisa kolom NIK, bukan spasi nama


# ----------------------------------------------------------------------------
# Extract / Transform
# ----------------------------------------------------------------------------
def _words(page, cond):
    f = page.filter(lambda o: o.get('object_type') != 'char' or cond(o))
    return f.extract_words(x_tolerance=1.5, keep_blank_chars=False)


def parse_pdf(file_obj):
    """Parse satu PDF HRIS. Return (period 'YYYYMM', list[dict] per karyawan)."""
    out, period = [], None
    with pdfplumber.open(file_obj) as pdf:
        for pg in pdf.pages:
            m = re.search(r'Bulan\s*:\s*(\d{6})', pg.extract_text() or '')
            if m:
                period = m.group(1)
            chars = pg.chars

            # posisi header tanggal 1..31
            hw = _words(pg, lambda o: 70 < o['top'] < 76 and o['size'] < 8.6)
            day_x = {int(w['text']): (w['x0'] + w['x1']) / 2 for w in hw
                     if w['text'].isdigit() and 1 <= int(w['text']) <= 31 and X_DAY_MIN < w['x0'] < X_DAY_MAX}
            if len(day_x) < 31:       # halaman overflow (kolom lanjutan), lewati
                continue

            # anchor tiap karyawan = nomor urut (font 9) di kolom paling kiri
            tops = sorted({round(c['top'], 1) for c in chars
                           if c['x0'] < 30 and c['size'] > 8.8 and c['text'].isdigit() and c['top'] > 78})
            anchors = []
            for t in tops:
                if not anchors or t - anchors[-1] > 3:
                    anchors.append(t)

            small_l = _words(pg, lambda o: o['size'] < 8.6 and o['x1'] <= X_TOTAL_SPLIT)
            small_r = _words(pg, lambda o: o['size'] < 8.6 and o['x1'] > X_TOTAL_SPLIT)
            small = small_l + small_r
            by_x = sorted(chars, key=lambda c: c['x0'])

            for T in anchors:
                near = [c for c in by_x if abs(c['top'] - T) < 2 and c['size'] > 8.8]
                no = ''.join(c['text'] for c in near if c['x0'] < 30)
                nik = ''.join(c['text'] for c in near if 30 <= c['x0'] < X_NAME_MIN)
                name = ''
                started = False
                for c in near:
                    if not (X_NAME_MIN <= c['x0'] < X_NAME_MAX):
                        continue
                    if not c['text'].strip():
                        if started and c['x0'] >= X_NAME_JUNK_SPACE:
                            name += ' '
                        continue
                    started = True
                    name += c['text']
                name = re.sub(r'\s+', ' ', name).strip()

                line_a = [w for w in small if -2.5 <= w['top'] - T <= 4 and w['x0'] >= X_DAY_MIN]
                line_b = [w for w in small if 8 <= w['top'] - T <= 16 and w['x0'] >= X_DAY_MIN]

                codes = {}
                for w in line_a:
                    if w['x0'] < X_SUMMARY_MIN and not re.fullmatch(r'[\d.]+', w['text']):
                        d = min(day_x, key=lambda k: abs(day_x[k] - (w['x0'] + w['x1']) / 2))
                        codes[d] = w['text']
                svals = [float(w['text']) for w in sorted(
                    (w for w in line_a if w['x0'] >= X_SUMMARY_MIN), key=lambda w: w['x0'])]

                ot, ot_total = {}, None
                for w in line_b:
                    v = float(w['text'])
                    if w['x1'] > X_TOTAL_SPLIT:
                        ot_total = v
                    else:
                        d = min(day_x, key=lambda k: abs(day_x[k] - (w['x0'] + w['x1']) / 2))
                        ot[d] = v
                out.append(dict(Period=period, No=int(no), NIK=nik, Name=name,
                                svals=svals, ot_total=ot_total, codes=codes, ot=ot))
    return period, out


def transform(records):
    """Raw records -> (daily df, monthly df, qa-issue list)."""
    daily, summ, issues = [], [], []
    for r in records:
        y, m = int(r['Period'][:4]), int(r['Period'][4:])
        nd = calendar.monthrange(y, m)[1]
        sv = dict(zip(SUM_COLS, r['svals']))

        # validasi 1: total lembur
        if round(sum(r['ot'].get(d, 0) for d in range(1, 32)) - (r['ot_total'] or 0), 2) != 0:
            issues.append((r['Period'], r['NIK'], 'Jumlah OT harian != total di PDF'))
        # validasi 2: hitung kode vs ringkasan PDF
        cc = Counter(RAW_ALIAS.get(v, v) for v in r['codes'].values())
        for code in ['H', 'A', 'I', 'CT', 'S', 'SD', 'TL', 'IR', 'DL', 'CH', 'PC', 'N']:
            if cc.get(code, 0) != sv.get(code, 0):
                issues.append((r['Period'], r['NIK'], f'Hitungan {code}: harian={cc.get(code, 0)} vs PDF={sv.get(code)}'))

        for d in range(1, nd + 1):
            date = dt.date(y, m, d)
            raw = r['codes'].get(d)
            code = RAW_ALIAS.get(raw, raw)
            source = 'PDF'
            if code is None:
                manual = MANUAL_OVERRIDES.get((r['NIK'], date.isoformat()))
                if manual:
                    code, source = manual, 'Manual'
                else:
                    code = 'OFF' if date.weekday() >= 5 else 'NODATA'
            daily.append([int(r['Period']), date, date.strftime('%a'), r['NIK'], r['Name'], raw, code,
                          STATUS.get(code, ('', 'Unknown'))[1],
                          0 if code in ('OFF', 'H', 'NODATA') else 1,
                          r['ot'].get(d, 0.0), source])
        summ.append([int(r['Period']), r['No'], r['NIK'], r['Name']] +
                    [sv.get(c) for c in SUM_COLS] + [r['ot_total']])

    D = pd.DataFrame(daily, columns=['Period', 'Date', 'Day_Name', 'NIK', 'Name', 'Raw_Code', 'Status',
                                     'Status_Group', 'Is_Workday', 'OT_Hours', 'Status_Source'])
    S = pd.DataFrame(summ, columns=['Period', 'No', 'NIK', 'Name'] + SUM_COLS + ['OT_Total_Reported'])
    return D, S, issues


def read_team_file(upload):
    """Baca file tim (CSV/XLSX) dengan kolom NAMA, NIK, TIM -> df[NIK, Team, Nama_File]."""
    name = upload.name.lower()
    raw = pd.read_csv(upload, dtype=str, sep=None, engine='python') if name.endswith(('.csv', '.txt')) \
        else pd.read_excel(upload, dtype=str)
    raw.columns = [str(c).strip().upper() for c in raw.columns]
    alias = {'NIK': ['NIK', 'ID', 'NIP'], 'TEAM': ['TIM', 'TEAM', 'TIM KERJA'], 'NAMA': ['NAMA', 'NAME']}
    pick = {}
    for key, opts in alias.items():
        pick[key] = next((c for c in raw.columns if c in opts), None)
    if not pick['NIK'] or not pick['TEAM']:
        raise ValueError(f'Kolom wajib NIK dan TIM tidak ditemukan. Kolom terbaca: {list(raw.columns)}')
    t = pd.DataFrame({
        'NIK': raw[pick['NIK']].astype(str).str.strip().str.upper(),
        'Team': raw[pick['TEAM']].astype(str).str.strip(),
        'Nama_File': raw[pick['NAMA']].astype(str).str.strip() if pick['NAMA'] else '',
    })
    t = t[(t.NIK != '') & (t.NIK != 'NAN')]
    t['Team'] = t.Team.replace({'': 'Unmapped', 'nan': 'Unmapped', 'NaN': 'Unmapped'})
    return t


def apply_team(D, S, team):
    """Tambah kolom Team ke D (paling kanan) dan S (setelah Name). NIK tanpa tim -> 'Unmapped'.
    Return D, S, dict info (duplikat, tidak ada di PDF, tidak ada di file tim, nama beda)."""
    dup = team[team.NIK.duplicated(keep=False)]
    team1 = team.drop_duplicates('NIK', keep='first').set_index('NIK')
    D = D.copy(); S = S.copy()
    D['Team'] = D.NIK.map(team1.Team).fillna('Unmapped')
    S.insert(S.columns.get_loc('Name') + 1, 'Team', S.NIK.map(team1.Team).fillna('Unmapped'))

    pdf_nik = D.drop_duplicates('NIK').set_index('NIK').Name
    norm = lambda x: re.sub(r'[^A-Z]', '', str(x).upper())
    mism = []
    for nik, nm in pdf_nik.items():
        if nik in team1.index and team1.loc[nik, 'Nama_File'] and norm(team1.loc[nik, 'Nama_File']) != norm(nm):
            mism.append((nik, nm, team1.loc[nik, 'Nama_File']))
    info = dict(
        duplicates=dup.sort_values('NIK'),
        unmapped=pdf_nik[~pdf_nik.index.isin(team1.index)].reset_index().rename(columns={'Name': 'Name_PDF'}),
        not_in_pdf=team1[~team1.index.isin(pdf_nik.index)].reset_index()[['NIK', 'Nama_File', 'Team']],
        name_mismatch=pd.DataFrame(mism, columns=['NIK', 'Name_PDF', 'Nama_File_Tim']),
    )
    return D, S, info


def team_summary(D):
    """Ringkasan per Team per Period."""
    w = D[D.Is_Workday == 1]
    g = w.groupby(['Team', 'Period'])
    out = pd.DataFrame({
        'Karyawan': D.groupby(['Team', 'Period']).NIK.nunique(),
        'Hari_Kerja': g.size(),
        'Hadir': g.apply(lambda x: (x.Status_Group == 'Hadir').sum()),
        'Tidak_Hadir': g.apply(lambda x: x.Status_Group.isin(['Cuti', 'Sakit', 'Izin', 'Alpa']).sum()),
        'Terlambat': g.apply(lambda x: (x.Status == 'TL').sum()),
        'Jam_Lembur': D.groupby(['Team', 'Period']).OT_Hours.sum(),
    }).reset_index()
    out['Tingkat_Hadir'] = (out.Hadir / out.Hari_Kerja).round(4)
    return out


def filter_inactive(D, S, max_missing=1):
    """Hapus karyawan yang dianggap RESIGN: tidak muncul di N periode terakhir.

    Dihitung dari periode terbaru yang diupload, mundur ke belakang:
      bulan_hilang_berturut = jumlah periode terbaru yang tidak ada datanya (dihitung dari akhir).
    - Jika karyawan masih ada di periode terbaru -> 0 -> dipertahankan (bukan resign).
    - Jika bulan_hilang_berturut > N -> seluruh barisnya dihapus dari D dan S.
    Bulan kosong di tengah (misal cuti panjang lalu kembali) tidak dihitung.
    Return (D_bersih, S_bersih, df_dihapus).
    """
    periods = sorted(D.Period.unique())
    pos = {p: i for i, p in enumerate(periods)}
    last = D.groupby('NIK').Period.max().map(pos)
    trailing = (len(periods) - 1 - last).rename('Bulan_Tidak_Ada_Terakhir')
    drop_nik = trailing[trailing > max_missing].index
    info = (D[D.NIK.isin(drop_nik)].groupby(['NIK', 'Name'])
            .agg(Periode_Terakhir=('Period', 'max'),
                 Periode_Ada=('Period', lambda x: ', '.join(map(str, sorted(x.unique())))))
            .reset_index())
    info['Bulan_Tidak_Ada_Terakhir'] = info.NIK.map(trailing)
    info['Alasan'] = f'Tidak ada di {int(max_missing) + 1}+ periode terakhir (periode terbaru {periods[-1]})'
    info = info[['NIK', 'Name', 'Periode_Terakhir', 'Bulan_Tidak_Ada_Terakhir', 'Periode_Ada', 'Alasan']]
    return D[~D.NIK.isin(drop_nik)].reset_index(drop=True), S[~S.NIK.isin(drop_nik)].reset_index(drop=True), info


# ----------------------------------------------------------------------------
# Load -> Excel
# ----------------------------------------------------------------------------
def build_workbook(D, S, removed=None):
    F = 'Arial'
    hfont = Font(name=F, bold=True, color='FFFFFF', size=10)
    hfill = PatternFill('solid', fgColor='1F3864')
    bfont = Font(name=F, size=10)

    def header(ws, n):
        for c in range(1, n + 1):
            cell = ws.cell(row=1, column=c)
            cell.font, cell.fill = hfont, hfill
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.freeze_panes = 'A2'

    def body(ws, n):
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=n):
            for c in row:
                c.font = bfont

    def widths(ws, w):
        for i, x in enumerate(w, 1):
            ws.column_dimensions[get_column_letter(i)].width = x

    wb = Workbook()
    ws = wb.active
    ws.title = 'README'
    periods = sorted(D.Period.unique())
    notes = [
        ('HRIS Monthly Attendance - hasil konsolidasi ETL', True),
        (f'Periode: {periods[0]} s/d {periods[-1]} | dibuat {dt.datetime.now():%Y-%m-%d %H:%M}', False),
        ('', False),
        ('Sheet', True),
        ('Attendance_Daily : 1 baris per karyawan per tanggal (tabel fakta untuk dashboard)', False),
        ('Employee_Monthly : ringkasan per karyawan per bulan (angka PDF) + kolom cek berformula', False),
        ('Dim_Employee / Dim_Status : tabel dimensi', False),
        ('Kolom Team (jika file tim diupload) : dipetakan dari NIK; NIK tanpa tim diberi Unmapped', False),
        ('Removed_Employees : karyawan yang dianggap resign (tidak ada di N periode terakhir) dan dihapus (jika filter aktif)', False),
        ('QA_Checks : rekap validasi', False),
        ('', False),
        ('Asumsi', True),
        ('1. Hanya hari kerja yang berkode di PDF; Sabtu/Minggu tanpa kode diisi Status = OFF.', False),
        ('2. Kode I* dinormalisasi ke IR (sesuai hitungan ringkasan HRIS); kode asli di Raw_Code.', False),
        ('3. Baris angka kedua per karyawan diasumsikan jam LEMBUR (OT_Hours).', False),
        ('3b. Koreksi manual (MANUAL_OVERRIDES di script) ditandai Status_Source = Manual; Raw_Code tetap kosong.', False),
        ('4. Hari kerja tanpa kode diberi Status NODATA (misal karyawan keluar/masuk di tengah bulan).', False),
        ('5. Kolom CHK, HD, LC, CL4, CL5, CL6, SOR tidak punya legenda di PDF; disimpan apa adanya.', False),
    ]
    for i, (t, b) in enumerate(notes, 1):
        ws.cell(row=i, column=1, value=t).font = Font(name=F, bold=b, size=12 if i == 1 else 10)
    ws.column_dimensions['A'].width = 120

    # Attendance_Daily
    ws = wb.create_sheet('Attendance_Daily')
    ws.append(list(D.columns))
    for r in D.itertuples(index=False):
        ws.append([None if (isinstance(v, float) and pd.isna(v)) else v for v in r])
    nd = ws.max_row
    header(ws, D.shape[1]); body(ws, D.shape[1])
    for r in range(2, nd + 1):
        ws.cell(row=r, column=2).number_format = 'yyyy-mm-dd'
        ws.cell(row=r, column=10).number_format = '0.0'
    widths(ws, [10, 12, 10, 10, 30, 10, 9, 13, 11, 10, 14] + ([16] if 'Team' in D.columns else []))
    ws.auto_filter.ref = f'A1:{get_column_letter(D.shape[1])}{nd}'

    # Employee_Monthly (+ formula check)
    ws = wb.create_sheet('Employee_Monthly')
    hdr = list(S.columns) + ['N_Calc', 'OT_Calc', 'Check_N', 'Check_OT']
    ws.append(hdr)
    ncs = S.shape[1]
    cN = get_column_letter(list(S.columns).index('N') + 1)
    cOT = get_column_letter(ncs)
    cNc, cOc = get_column_letter(ncs + 1), get_column_letter(ncs + 2)
    for i, r in enumerate(S.itertuples(index=False), 2):
        ws.append([None if (isinstance(v, float) and pd.isna(v)) else v for v in r] + [
            f'=COUNTIFS(Attendance_Daily!$A:$A,$A{i},Attendance_Daily!$D:$D,$C{i},Attendance_Daily!$G:$G,"N")',
            f'=SUMIFS(Attendance_Daily!$J:$J,Attendance_Daily!$A:$A,$A{i},Attendance_Daily!$D:$D,$C{i})',
            f'=IF({cNc}{i}={cN}{i},"OK","DIFF")',
            f'=IF(ROUND({cOc}{i}-{cOT}{i},2)=0,"OK","DIFF")'])
    ns = ws.max_row
    header(ws, len(hdr)); body(ws, len(hdr))
    widths(ws, [10, 6, 10, 30] + ([16] if 'Team' in S.columns else []) + [6] * len(SUM_COLS) + [12, 9, 9, 9, 9])
    ws.auto_filter.ref = f'A1:{get_column_letter(len(hdr))}{ns}'
    ws.row_dimensions[1].height = 30

    # Dim_Employee
    ws = wb.create_sheet('Dim_Employee')
    ecols = ['NIK', 'Name'] + (['Team'] if 'Team' in D.columns else [])
    ws.append(ecols + ['Periods_Count'])
    for e in D.groupby(ecols).Period.nunique().reset_index().itertuples(index=False):
        ws.append(list(e))
    header(ws, len(ecols) + 1); body(ws, len(ecols) + 1); widths(ws, [10, 30, 16, 14][:len(ecols) + 1] if len(ecols) == 3 else [10, 30, 14])
    ne = ws.max_row

    # Dim_Status
    ws = wb.create_sheet('Dim_Status')
    ws.append(['Status', 'Description', 'Status_Group'])
    for k, (d, g) in STATUS.items():
        ws.append([k, d, g])
    header(ws, 3); body(ws, 3); widths(ws, [10, 34, 14])

    # QA_Checks
    ws = wb.create_sheet('QA_Checks')
    ws.append(['Check', 'Result', 'Expected'])
    cCN, cCO = get_column_letter(ncs + 3), get_column_letter(ncs + 4)
    for q in [
        ['Baris Attendance_Daily', f'=COUNTA(Attendance_Daily!$A$2:$A${nd})', len(D)],
        ['Baris Employee_Monthly', f'=COUNTA(Employee_Monthly!$A$2:$A${ns})', len(S)],
        ['Karyawan unik', f'=COUNTA(Dim_Employee!$A$2:$A${ne})', D.NIK.nunique()],
        ['Check_N = DIFF', f'=COUNTIF(Employee_Monthly!${cCN}$2:${cCN}${ns},"DIFF")', 0],
        ['Check_OT = DIFF', f'=COUNTIF(Employee_Monthly!${cCO}$2:${cCO}${ns},"DIFF")', 0],
        ['Hari kerja NODATA', f'=COUNTIF(Attendance_Daily!$G$2:$G${nd},"NODATA")', int((D.Status == 'NODATA').sum())],
        ['Hari dikoreksi manual (Status_Source = Manual)', f'=COUNTIF(Attendance_Daily!$K$2:$K${nd},"Manual")', int((D.Status_Source == 'Manual').sum())],
    ] + ([['Baris tanpa tim (Unmapped)', f'=COUNTIF(Attendance_Daily!${get_column_letter(D.columns.get_loc("Team") + 1)}$2:${get_column_letter(D.columns.get_loc("Team") + 1)}${nd},"Unmapped")', int((D.Team == 'Unmapped').sum())]] if 'Team' in D.columns else []):
        ws.append(q)
    header(ws, 3); body(ws, 3); widths(ws, [30, 14, 14])

    if removed is not None and len(removed):
        ws = wb.create_sheet('Removed_Employees')
        ws.append(list(removed.columns))
        for r in removed.itertuples(index=False):
            ws.append(list(r))
        header(ws, removed.shape[1]); body(ws, removed.shape[1]); widths(ws, [10, 30, 16, 14, 40, 50])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# Streamlit UI
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def run_etl(files):
    """files: tuple[(filename, bytes)] -> (D, S, issues, per_file_info)"""
    recs, info = [], []
    for fname, content in files:
        period, rows = parse_pdf(io.BytesIO(content))
        info.append(dict(File=fname, Period=period, Employees=len(rows)))
        recs += rows
    D, S, issues = transform(recs)
    return D, S, issues, info


def main():
    st.set_page_config(page_title='HRIS Attendance ETL', layout='wide')
    st.title('HRIS Attendance ETL')
    st.caption('PDF Monthly Attendance Report (PT.GLOSTAR INDONESIA) -> Excel siap dashboard')

    with st.sidebar:
        st.header('Master data tim')
        team_up = st.file_uploader('Upload master tim (kolom: NAMA, NIK, TIM) - CSV / XLSX',
                                   type=['csv', 'xlsx', 'xls'], key='team_file')
        st.caption('Opsional. Template berisi NAMA dan NIK dari PDF tersedia di sini setelah PDF diproses.')
        tmpl_slot = st.container()

    up = st.file_uploader('Upload PDF HRIS (boleh banyak bulan sekaligus)', type='pdf', accept_multiple_files=True)
    if not up:
        st.session_state.pop('etl', None)   # reset hasil lama kalau semua file dihapus
        st.info('Upload satu atau beberapa PDF untuk memulai.')
        return

    files = tuple((f.name, f.getvalue()) for f in up)
    sig = tuple((n, len(c)) for n, c in files)   # penanda untuk mendeteksi file berubah

    # ---- Tombol proses ----
    if st.button('Proses', type='primary'):
        try:
            with st.spinner('Memproses PDF...'):
                D, S, issues, info = run_etl(files)
            st.session_state['etl'] = dict(D=D, S=S, issues=issues, info=info, sig=sig)
        except Exception as e:
            st.session_state.pop('etl', None)
            st.error(f'Gagal memproses PDF: {e}')
            return

    etl = st.session_state.get('etl')
    if etl is None:
        st.info('Klik tombol **Proses** untuk mulai membaca PDF.')
        return
    if etl['sig'] != sig:
        st.warning('Daftar file berubah sejak terakhir diproses. Klik **Proses** lagi untuk memperbarui hasil.')

    D, S, issues, info = etl['D'], etl['S'], etl['issues'], etl['info']
    if D.empty:
        st.warning('Tidak ada data karyawan yang terbaca. Pastikan format PDF sama dengan laporan HRIS.')
        return

    # Mapping tim (opsional) - uploader ada di sidebar
    st.subheader('Mapping tim')
    tmpl = D.drop_duplicates('NIK')[['Name', 'NIK']].rename(columns={'Name': 'NAMA'}).assign(TIM='')
    tmpl_slot.download_button('Download template tim (CSV)', tmpl.to_csv(index=False).encode('utf-8'),
                              'template_tim.csv', 'text/csv')
    if team_up is None:
        st.info('Belum ada master tim. Upload lewat sidebar kiri (Master data tim) untuk menambahkan kolom Team.')
    team_info = None
    if team_up is not None:
        try:
            D, S, team_info = apply_team(D, S, read_team_file(team_up))
        except Exception as e:
            st.error(f'Gagal membaca file tim: {e}')
            team_info = None
        if team_info is not None:
            n_un = len(team_info['unmapped'])
            if n_un:
                st.warning(f'{n_un} karyawan di PDF tidak ada di file tim (Team = Unmapped):')
                st.dataframe(team_info['unmapped'], use_container_width=True)
            else:
                st.success('Semua karyawan di PDF sudah punya tim.')
            if len(team_info['duplicates']):
                st.warning('NIK ganda di file tim (dipakai baris pertama):')
                st.dataframe(team_info['duplicates'], use_container_width=True)
            if len(team_info['not_in_pdf']):
                st.info(f"{len(team_info['not_in_pdf'])} NIK ada di file tim tapi tidak ada di PDF.")
                with st.expander('Lihat daftar'):
                    st.dataframe(team_info['not_in_pdf'], use_container_width=True)
            if len(team_info['name_mismatch']):
                st.info(f"{len(team_info['name_mismatch'])} NIK namanya berbeda antara PDF dan file tim (cek salah ketik / NIK tertukar).")
                with st.expander('Lihat daftar'):
                    st.dataframe(team_info['name_mismatch'], use_container_width=True)

    # Filter karyawan yang tidak ada lebih dari N bulan
    f1, f2 = st.columns([2, 1])
    use_filter = f1.checkbox('Hapus karyawan resign (tidak ada di N periode terakhir)', value=True)
    max_missing = f2.number_input('N (bulan)', min_value=0, max_value=24, value=1, step=1)
    removed = pd.DataFrame()
    if use_filter and D.Period.nunique() > 1:
        D, S, removed = filter_inactive(D, S, int(max_missing))
        if len(removed):
            st.warning(f'{len(removed)} karyawan dihapus (tidak ada di {int(max_missing) + 1}+ periode terakhir, dianggap resign):')
            st.dataframe(removed, use_container_width=True)
        else:
            st.caption(f'Semua karyawan masih ada dalam {int(max_missing) + 1} periode terakhir.')

    dup = pd.DataFrame(info).Period.duplicated().any()
    if dup:
        st.warning('Ada lebih dari satu file dengan periode yang sama. Data akan terduplikasi.')

    # KPI
    w = D[D.Is_Workday == 1]
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric('Periode', D.Period.nunique())
    c2.metric('Karyawan unik', D.NIK.nunique())
    c3.metric('Baris harian', f'{len(D):,}')
    c4.metric('Tingkat hadir', f'{(w.Status_Group == "Hadir").mean():.1%}' if len(w) else '-')
    c5.metric('Total jam lembur', f'{D.OT_Hours.sum():,.1f}')

    # QA
    if issues:
        st.error(f'{len(issues)} selisih validasi terhadap angka ringkasan PDF')
        st.dataframe(pd.DataFrame(issues, columns=['Period', 'NIK', 'Issue']), use_container_width=True)
    else:
        st.success('Validasi lulus: hitungan kode status dan total lembur cocok dengan ringkasan PDF.')
    nodata = D[D.Status == 'NODATA']
    if len(nodata):
        rpt = (nodata.groupby(['Period', 'NIK', 'Name'])
               .agg(Hari=('Date', 'count'),
                    Tanggal=('Date', lambda x: ', '.join(d.strftime('%d') for d in sorted(x))))
               .reset_index())
        st.warning(f'LAPORAN: {len(nodata)} hari kerja kosong di PDF (NODATA) pada {rpt.NIK.nunique()} karyawan. '
                   'Cek ke HRD; kalau sudah dikonfirmasi tambahkan ke MANUAL_OVERRIDES.')
        st.dataframe(rpt, use_container_width=True)
    n_manual = int((D.Status_Source == 'Manual').sum())
    if n_manual:
        st.caption(f'{n_manual} hari dikoreksi manual (Status_Source = Manual).')

    t1, t2, t3, t4, t5 = st.tabs(['Attendance_Daily', 'Employee_Monthly', 'Ringkasan status', 'Ringkasan tim', 'File terbaca'])
    with t1:
        per = st.multiselect('Filter periode', sorted(D.Period.unique()), default=sorted(D.Period.unique()))
        view = D[D.Period.isin(per)]
        if 'Team' in D.columns:
            tms = st.multiselect('Filter tim', sorted(D.Team.unique()), default=sorted(D.Team.unique()))
            view = view[view.Team.isin(tms)]
        st.dataframe(view, use_container_width=True, height=400)
    with t2:
        st.dataframe(S, use_container_width=True, height=400)
    with t3:
        piv = w.pivot_table(index='Period', columns='Status', values='NIK', aggfunc='count', fill_value=0)
        st.dataframe(piv, use_container_width=True)
        st.bar_chart(w.pivot_table(index='Period', columns='Status_Group', values='NIK',
                                   aggfunc='count', fill_value=0))
    with t4:
        if 'Team' in D.columns:
            ts = team_summary(D)
            st.dataframe(ts, use_container_width=True, height=400,
                         column_config={'Tingkat_Hadir': st.column_config.NumberColumn(format='percent')})
            st.bar_chart(ts.pivot_table(index='Team', columns='Period', values='Tingkat_Hadir'))
        else:
            st.info('Upload file tim untuk melihat ringkasan per tim.')
    with t5:
        st.dataframe(pd.DataFrame(info), use_container_width=True)

    st.divider()
    xlsx = build_workbook(D, S, removed)
    stamp = f'{D.Period.min()}-{D.Period.max()}'
    d1, d2 = st.columns(2)
    d1.download_button('Download Excel (semua sheet)', xlsx, f'HRIS_Attendance_{stamp}.xlsx',
                       'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', type='primary')
    d2.download_button('Download CSV (Attendance_Daily, untuk Looker Studio)',
                       D.to_csv(index=False).encode('utf-8'), f'Attendance_Daily_{stamp}.csv', 'text/csv')


if __name__ == '__main__':
    main()
