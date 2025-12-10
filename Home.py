"""Home page for PGD Apps — simple layout with system theme support.
This file uses CSS variables + prefers-color-scheme for light/dark theme,
no sidebar theme selector.
"""
import streamlit as st
from utils import set_page, header, footer
from pathlib import Path

set_page("PGD Apps — Home", "🤖")

# Theme variables (light + dark)
light_vars = """
--primary-color: #1f77b4;
--primary-dark: #0d47a1;
--accent-color: #ffc107;
--text-primary: #111827;
--text-secondary: #4b5563;
--bg-primary: #ffffff;
--bg-secondary: #f8fafc;
--card-bg: #ffffff;
--border-color: #e6eef8;
--shadow: rgba(16,24,40,0.06);
"""

dark_vars = """
--primary-color: #5b9ee1;
--primary-dark: #7fb3f0;
--accent-color: #ffa500;
--text-primary: #e6eef8;
--text-secondary: #a0aec0;
--bg-primary: #0b1220;
--bg-secondary: #071123;
--card-bg: #0f1724;
--border-color: #1f2937;
--shadow: rgba(2,6,23,0.8);
"""

# Common CSS uses the variables above
common_css = """
:root { color-scheme: light; }

.stApp, .main, body {
    background: var(--bg-primary) !important;
    color: var(--text-primary) !important;
}

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, var(--bg-secondary), var(--bg-primary)) !important;
    color: var(--text-primary) !important;
    border-right: 1px solid var(--border-color);
}

.hero-section {
    background: linear-gradient(135deg, var(--primary-color) 0%, var(--primary-dark) 100%);
    color: white;
    padding: 3rem 2rem;
    border-radius: 10px;
    margin-bottom: 2rem;
    text-align: center;
    box-shadow: 0 4px 12px var(--shadow);
    border: 1px solid rgba(255,255,255,0.06);
    transition: all 0.25s ease;
}

.tool-card {
    background: linear-gradient(135deg, var(--card-bg) 0%, var(--bg-primary) 100%);
    border: 1px solid var(--border-color);
    border-radius: 8px;
    padding: 1.25rem;
    margin: 1rem 0;
    transition: transform 0.18s ease, box-shadow 0.18s ease, border-color 0.18s ease;
    box-shadow: 0 2px 6px var(--shadow);
    color: var(--text-primary);
}
.tool-card:hover { transform: translateY(-4px); box-shadow: 0 10px 30px rgba(59,130,246,0.08); border-color: var(--primary-color); }
.tool-title { color: var(--primary-color); font-weight: 600; margin-bottom: 0.35rem; }
.tool-description { color: var(--text-secondary); line-height: 1.45; }

.stat-box { text-align:center; min-width:150px; padding:1rem; background: linear-gradient(135deg,var(--card-bg) 0%,var(--bg-primary) 100%); border-radius:8px; margin:0.5rem; border:1px solid var(--border-color); box-shadow:0 2px 8px var(--shadow); }
.stat-number { font-size:1.8rem; color:var(--primary-color); font-weight:700; }
.stat-label { color:var(--text-secondary); font-size:0.9rem; }

.info-banner { background: linear-gradient(135deg, rgba(255,243,205,0.85), rgba(255,250,240,0.7)); border-left:4px solid var(--accent-color); padding:1rem; border-radius:6px; margin:1.25rem 0; color:var(--text-primary); }
.info-banner-info { background: linear-gradient(135deg, rgba(209,236,241,0.85), rgba(232,244,248,0.6)); border-left-color: #17a2b8; color:var(--text-primary); }

.section-divider { border:none; height:2px; background: linear-gradient(to right, var(--primary-color), transparent); margin:2rem 0; opacity:0.45; }

/* Buttons fallback */
.stButton>button { background: linear-gradient(135deg, var(--primary-color), var(--primary-dark)); color: white; border-radius:8px; }

/* Dataframe */
.stDataFrame, table.dataframe { background: var(--card-bg) !important; color: var(--text-primary) !important; }

* { transition: background-color 0.2s ease, color 0.2s ease, border-color 0.2s ease; }

@media (max-width:768px) {
  .hero-section{ padding:1.5rem 1rem; }
  .hero-section h1{ font-size:1.6rem; }
}
"""

# Build CSS using prefers-color-scheme so app follows OS theme
css = f"""
<style>
:root {{
    {light_vars}
}}
@media (prefers-color-scheme: dark) {{
    :root {{
        {dark_vars}
    }}
    :root {{ color-scheme: dark; }}
}}
{common_css}
</style>
"""

st.markdown(css, unsafe_allow_html=True)

header("🤖 PGD Apps — Home")

# Hero Section
st.markdown("""
<div class="hero-section">
    <h1>Selamat Datang di PGD Apps</h1>
    <p>Kumpulan tools otomasi harian untuk tim PGD yang powerful dan user-friendly</p>
</div>
""", unsafe_allow_html=True)

# Logo dan intro
col1, col2 = st.columns([1, 3])
with col1:
    logo = Path(__file__).resolve().parent / "assets" / "logo.png"
    if logo.exists():
        st.image(str(logo), width=140, caption="PGD Logo")
    else:
        st.info("Logo tidak ditemukan di assets/")

with col2:
    st.markdown("""
    ### Tentang Aplikasi Ini
    PGD Apps adalah platform terintegrasi yang dirancang untuk meningkatkan efisiensi 
    kerja tim PGD melalui otomasi proses-proses penting. Dengan antarmuka yang intuitif, 
    tools ini membantu Anda mengelola data, membuat laporan, dan menganalisis informasi 
    dengan lebih cepat.

    **✨ Info:** Tampilan mengikuti preferensi sistem (Light/Dark). Untuk mengubah, atur tema OS Anda.
    """)

# Stats Section
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("📊 Tools Tersedia", 12)
with col2:
    st.metric("⚡ Fitur Utama", 50)
with col3:
    st.metric("📈 Efisiensi", "40%+")
with col4:
    st.metric("🌙 Tema", "System")

# Tools Section
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
st.subheader("📚 Daftar Tools & Fitur")

pages_dir = Path(__file__).resolve().parent / "pages"
if pages_dir.exists():
    items = sorted(pages_dir.glob("*.py"))
    descriptions = {
        "1_Quantity Change Extractor": "🧾 Ekstrak dan normalisasi data quantity, reshape UK_*, dan bandingkan perubahan qty dengan akurat.",
        "2_Input PGD_WFM_BTP_Tracking_Report": "📋 Input dan kelola tracking report PGD WFM BTP dengan sistem yang terstruktur.",
        "3_Merger Daily Report": "📦 Rekap dan merger banyak file laporan harian menjadi satu output yang rapi.",
        "4_Jadwal Audit": "📅 Generator jadwal audit mingguan/bulanan dengan format siap pakai dan exportable.",
        "5_Reroute Tools": "🔄 Bandingkan Old vs New PO, cek konsistensi size, dan batch PO Finder otomatis.",
        "6_Input Tracking Report Pending Cancel": "⏳ Tracking status pekerjaan pending/cancel dan ekspor hasilnya ke berbagai format.",
        "7_Susun Sizelist": "📏 Susun dan standarkan daftar size sesuai kebutuhan produksi dengan validasi otomatis.",
        "8_Tooling Sizelist": "🔧 Kelola sizelist tooling dengan fitur import, export, dan validasi data lengkap.",
        "9_Rekap E-Memo": "📧 Rekap data dari email/memo dan buatkan laporan terintegrasi dengan mudah.",
        "10_Check Export Plan Daily and Monthly": "🌍 Check SO rencana export harian vs bulanan dengan identifikasi kasus yang ada.",
        "11_Comparison RSA": "🔍 Analisis dan bandingkan data RSA dengan statistik performa yang detail.",
        "12_Advanced_Analytics": "📊 Advanced Analytics Dashboard dengan visualisasi data dan statistik mendalam.",
    }
    for p in items:
        name = p.stem
        desc = descriptions.get(name, f"Tool khusus untuk: {name}")
        st.markdown(f"""
        <div class="tool-card" role="group" aria-label="{name}">
            <div class="tool-title">{desc.split()[0]} {' '.join(desc.split()[1:])}</div>
            <div class="tool-description">{desc}</div>
        </div>
        """, unsafe_allow_html=True)

# Info Section
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
col1, col2 = st.columns([2, 1])
with col1:
    st.markdown("""
    <div class="info-banner" role="note">
    <strong>💡 Tips Penggunaan</strong><br>
    • Gunakan sidebar untuk navigasi antar tool<br>
    • Setiap tool memiliki bantuan interaktif, hover/klik untuk melihat detail<br>
    • Selalu backup data sebelum melakukan proses besar<br>
    • Untuk hasil optimal, ikuti instruksi sesuai format yang diminta
    </div>
    """, unsafe_allow_html=True)
with col2:
    st.markdown("""
    <div class="info-banner info-banner-info" role="note">
    <strong>📞 Dukungan</strong><br>
    <small>Ada pertanyaan atau bug?<br>
    📧 nazarudin@gsid.co.id<br>
    Respon cepat & helpful support</small>
    </div>
    """, unsafe_allow_html=True)

footer("PGD Apps • Made by Nazarudin Zaini :D")