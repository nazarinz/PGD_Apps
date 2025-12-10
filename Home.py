"""Home page for PGD Apps — simple layout with system theme support.
This file uses CSS variables + prefers-color-scheme for light/dark theme,
no sidebar theme selector.
"""
import streamlit as st
from utils import set_page, header, footer
from pathlib import Path

set_page("PGD Apps — Home", "🤖")

# Inject JavaScript to force theme application
st.markdown("""
<script>
    // Force apply theme colors
    const applyTheme = () => {
        const isDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
        const root = document.documentElement;
        const stApp = document.querySelector('.stApp');
        const main = document.querySelector('.main');
        
        if (isDark) {
            root.style.setProperty('background-color', '#0f172a', 'important');
            if (stApp) stApp.style.backgroundColor = '#0f172a';
            if (main) main.style.backgroundColor = '#0f172a';
        } else {
            root.style.setProperty('background-color', '#ffffff', 'important');
            if (stApp) stApp.style.backgroundColor = '#ffffff';
            if (main) main.style.backgroundColor = '#ffffff';
        }
    };
    
    // Apply on load and theme change
    setTimeout(applyTheme, 100);
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', applyTheme);
</script>
""", unsafe_allow_html=True)

# Complete CSS with forced backgrounds
css = """
<style>
/* Force light mode as default */
:root {
    --primary-color: #2563eb;
    --primary-dark: #1d4ed8;
    --accent-color: #0ea5e9;
    --text-primary: #0f172a;
    --text-secondary: #475569;
    --bg-primary: #ffffff;
    --bg-secondary: #f8fafc;
    --card-bg: #fafbff;
    --border-color: #e0e7ff;
    --shadow: rgba(37, 99, 235, 0.08);
    --hero-gradient-start: #3b82f6;
    --hero-gradient-end: #2563eb;
    color-scheme: light;
}

/* Dark mode override */
@media (prefers-color-scheme: dark) {
    :root {
        --primary-color: #60a5fa;
        --primary-dark: #3b82f6;
        --accent-color: #93c5fd;
        --text-primary: #f1f5f9;
        --text-secondary: #cbd5e1;
        --bg-primary: #0f172a;
        --bg-secondary: #1e293b;
        --card-bg: #1e293b;
        --border-color: #334155;
        --shadow: rgba(96, 165, 250, 0.15);
        --hero-gradient-start: #1e40af;
        --hero-gradient-end: #3730a3;
        color-scheme: dark;
    }
}

/* Aggressive background forcing */
html, body {
    background-color: var(--bg-primary) !important;
    background: var(--bg-primary) !important;
}

.stApp {
    background-color: var(--bg-primary) !important;
    background: var(--bg-primary) !important;
}

section[data-testid="stAppViewContainer"] {
    background-color: var(--bg-primary) !important;
    background: var(--bg-primary) !important;
}

section[data-testid="stAppViewContainer"] > .main {
    background-color: var(--bg-primary) !important;
    background: var(--bg-primary) !important;
}

.main .block-container {
    background-color: var(--bg-primary) !important;
}

.stApp, .main, body {
    color: var(--text-primary) !important;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: var(--bg-secondary) !important;
    background-color: var(--bg-secondary) !important;
    color: var(--text-primary) !important;
    border-right: 1px solid var(--border-color);
}

[data-testid="stSidebar"] > div:first-child {
    background-color: var(--bg-secondary) !important;
}

[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
    color: var(--text-primary) !important;
}

/* Hero Section */
.hero-section {
    background: linear-gradient(135deg, var(--hero-gradient-start) 0%, var(--hero-gradient-end) 100%);
    color: white;
    padding: 3.5rem 2rem;
    border-radius: 16px;
    margin-bottom: 2.5rem;
    text-align: center;
    box-shadow: 0 8px 32px var(--shadow);
    border: 1px solid rgba(255,255,255,0.1);
    transition: all 0.3s ease;
}

.hero-section h1 {
    font-size: 2.5rem;
    font-weight: 700;
    margin-bottom: 1rem;
    text-shadow: 0 2px 4px rgba(0,0,0,0.1);
    color: white !important;
}

.hero-section p {
    font-size: 1.15rem;
    opacity: 0.95;
    font-weight: 400;
    color: white !important;
}

/* Tool Cards */
.tool-card {
    background: var(--card-bg);
    border: 1.5px solid var(--border-color);
    border-radius: 12px;
    padding: 1.5rem;
    margin: 1rem 0;
    transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease;
    box-shadow: 0 1px 3px rgba(37, 99, 235, 0.06), 0 4px 8px rgba(37, 99, 235, 0.04);
    color: var(--text-primary);
    position: relative;
    overflow: hidden;
}

.tool-card::before {
    content: '';
    position: absolute;
    top: 0;
    left: 0;
    width: 4px;
    height: 100%;
    background: var(--primary-color);
    opacity: 0;
    transition: opacity 0.2s ease;
}

.tool-card:hover {
    transform: translateY(-4px);
    box-shadow: 0 4px 12px rgba(37, 99, 235, 0.12), 0 12px 32px rgba(37, 99, 235, 0.08);
    border-color: var(--primary-color);
}

.tool-card:hover::before {
    opacity: 1;
}

.tool-title {
    color: var(--primary-color);
    font-weight: 600;
    font-size: 1.1rem;
    margin-bottom: 0.5rem;
}

.tool-description {
    color: var(--text-secondary);
    line-height: 1.6;
    font-size: 0.95rem;
}

/* Stats */
.stat-box {
    text-align: center;
    padding: 1.5rem 1rem;
    background: var(--card-bg);
    border-radius: 12px;
    border: 1.5px solid var(--border-color);
    box-shadow: 0 1px 3px rgba(37, 99, 235, 0.06), 0 4px 8px rgba(37, 99, 235, 0.04);
    transition: transform 0.2s ease, box-shadow 0.2s ease;
}

.stat-box:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(37, 99, 235, 0.12), 0 8px 20px rgba(37, 99, 235, 0.08);
}

.stat-number {
    font-size: 2rem;
    color: var(--primary-color);
    font-weight: 700;
    margin-bottom: 0.25rem;
}

.stat-label {
    color: var(--text-secondary);
    font-size: 0.9rem;
    font-weight: 500;
}

/* Info Banner */
.info-banner {
    background: var(--card-bg);
    border-left: 4px solid var(--accent-color);
    padding: 1.25rem;
    border-radius: 8px;
    margin: 1.5rem 0;
    color: var(--text-primary);
    box-shadow: 0 2px 8px var(--shadow);
}

.info-banner strong {
    color: var(--primary-color);
    font-size: 1.05rem;
}

.info-banner-info {
    border-left-color: #06b6d4;
}

/* Divider */
.section-divider {
    border: none;
    height: 2px;
    background: linear-gradient(to right, var(--primary-color), transparent);
    margin: 2.5rem 0;
    opacity: 0.3;
}

/* Buttons */
.stButton>button {
    background: linear-gradient(135deg, var(--primary-color), var(--primary-dark));
    color: white;
    border-radius: 8px;
    font-weight: 600;
    transition: all 0.2s ease;
}

.stButton>button:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px var(--shadow);
}

/* Override Streamlit default text colors */
.stMarkdown, p, h1, h2, h3, h4, h5, h6, span, div {
    color: var(--text-primary) !important;
}

.stDataFrame, table.dataframe {
    background: var(--card-bg) !important;
    color: var(--text-primary) !important;
}

/* Smooth transitions */
* {
    transition: background-color 0.2s ease, color 0.2s ease, border-color 0.2s ease;
}

/* Mobile responsive */
@media (max-width: 768px) {
    .hero-section {
        padding: 2rem 1rem;
    }
    .hero-section h1 {
        font-size: 1.75rem;
    }
    .hero-section p {
        font-size: 1rem;
    }
}
</style>
"""

st.markdown(css, unsafe_allow_html=True)

header("🤖 PGD Apps — Home")

# Hero Section
st.markdown("""
<div class="hero-section">
    <h1>🚀 Selamat Datang di PGD Apps</h1>
    <p>Kumpulan tools otomasi harian user-friendly</p>
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
    ### 💼 Tentang Aplikasi Ini
    PGD Apps adalah Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor.

    **✨ Info:** Tampilan mengikuti preferensi sistem (Light/Dark). Untuk mengubah, atur tema OS Anda.
    """)

# Stats Section
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.markdown("""
    <div class="stat-box">
        <div class="stat-number">12</div>
        <div class="stat-label">📊 Tools Tersedia</div>
    </div>
    """, unsafe_allow_html=True)
with col2:
    st.markdown("""
    <div class="stat-box">
        <div class="stat-number">50+</div>
        <div class="stat-label">⚡ Fitur Utama</div>
    </div>
    """, unsafe_allow_html=True)
with col3:
    st.markdown("""
    <div class="stat-box">
        <div class="stat-number">40%</div>
        <div class="stat-label">📈 Efisiensi</div>
    </div>
    """, unsafe_allow_html=True)
with col4:
    st.markdown("""
    <div class="stat-box">
        <div class="stat-number">Auto</div>
        <div class="stat-label">🌙 Tema System</div>
    </div>
    """, unsafe_allow_html=True)

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
