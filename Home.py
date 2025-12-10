# Home.py
import streamlit as st
from utils import set_page, header, footer
from pathlib import Path

set_page("PGD Apps — Home", "🤖")

# Custom CSS untuk home page yang lebih menarik + DARK MODE SUPPORT
st.markdown("""
<style>
:root {
    --primary-color: #1f77b4;
    --primary-dark: #0d47a1;
    --accent-color: #ffc107;
    --text-light: #333;
    --text-dark: #e0e0e0;
    --bg-light: #ffffff;
    --bg-dark: #1e1e1e;
    --card-light: #f8f9fa;
    --card-dark: #2d2d2d;
    --border-light: #e0e0e0;
    --border-dark: #404040;
}

@media (prefers-color-scheme: dark) {
    :root {
        --primary-color: #5b9ee1;
        --primary-dark: #7fb3f0;
        --accent-color: #ffa500;
        --text-light: #e0e0e0;
        --text-dark: #a0a0a0;
        --bg-light: #1e1e1e;
        --bg-dark: #2d2d2d;
        --card-light: #2d2d2d;
        --card-dark: #353535;
        --border-light: #404040;
        --border-dark: #505050;
    }
}

.hero-section {
    background: linear-gradient(135deg, var(--primary-color) 0%, var(--primary-dark) 100%);
    color: white;
    padding: 3rem 2rem;
    border-radius: 10px;
    margin-bottom: 2rem;
    text-align: center;
    box-shadow: 0 4px 12px rgba(31, 119, 180, 0.2);
    border: 1px solid rgba(255, 255, 255, 0.1);
    transition: all 0.3s ease;
}

.hero-section:hover {
    box-shadow: 0 6px 16px rgba(31, 119, 180, 0.3);
}

.hero-section h1 {
    color: white;
    border: none;
    margin: 0;
    font-size: 2.5rem;
    text-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
}

.hero-section p {
    font-size: 1.1rem;
    margin: 0.5rem 0 0 0;
    opacity: 0.95;
    text-shadow: 0 1px 2px rgba(0, 0, 0, 0.1);
}

.tool-card {
    background: linear-gradient(135deg, var(--card-light) 0%, var(--bg-light) 100%);
    border: 1px solid var(--border-light);
    border-radius: 8px;
    padding: 1.5rem;
    margin: 1rem 0;
    transition: all 0.3s ease;
    box-shadow: 0 2px 6px rgba(0, 0, 0, 0.08);
    color: var(--text-light);
}

@media (prefers-color-scheme: dark) {
    .tool-card {
        background: linear-gradient(135deg, var(--card-light) 0%, var(--card-dark) 100%);
        border-color: var(--border-light);
    }
}

.tool-card:hover {
    transform: translateY(-4px);
    box-shadow: 0 8px 20px rgba(31, 119, 180, 0.15);
    border-color: var(--primary-color);
}

.tool-icon {
    font-size: 1.8rem;
    margin-right: 0.5rem;
}

.tool-title {
    color: var(--primary-color);
    font-size: 1.1rem;
    font-weight: 600;
    margin-bottom: 0.5rem;
}

.tool-description {
    color: var(--text-dark);
    font-size: 0.95rem;
    line-height: 1.5;
}

.stats-section {
    display: flex;
    justify-content: space-around;
    margin: 2rem 0;
    flex-wrap: wrap;
}

.stat-box {
    text-align: center;
    flex: 1;
    min-width: 150px;
    padding: 1rem;
    background: linear-gradient(135deg, var(--card-light) 0%, var(--bg-light) 100%);
    border-radius: 8px;
    margin: 0.5rem;
    border: 1px solid var(--border-light);
    transition: all 0.3s ease;
}

@media (prefers-color-scheme: dark) {
    .stat-box {
        background: linear-gradient(135deg, var(--card-light) 0%, var(--card-dark) 100%);
        border-color: var(--border-light);
    }
}

.stat-box:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(31, 119, 180, 0.1);
}

.stat-number {
    font-size: 2rem;
    color: var(--primary-color);
    font-weight: bold;
}

.stat-label {
    color: var(--text-dark);
    font-size: 0.9rem;
    margin-top: 0.5rem;
}

.info-banner {
    background: linear-gradient(135deg, #fff3cd 0%, #fffbea 100%);
    border-left: 4px solid #ffc107;
    padding: 1rem;
    border-radius: 6px;
    margin: 2rem 0;
    color: #333;
    transition: all 0.3s ease;
}

@media (prefers-color-scheme: dark) {
    .info-banner {
        background: linear-gradient(135deg, rgba(255, 193, 7, 0.15) 0%, rgba(255, 193, 7, 0.1) 100%);
        border-left-color: #ffd43b;
        color: #ffe066;
    }
}

.info-banner-info {
    background: linear-gradient(135deg, #d1ecf1 0%, #e8f4f8 100%);
    border-left-color: #17a2b8;
    color: #0c5460;
}

@media (prefers-color-scheme: dark) {
    .info-banner-info {
        background: linear-gradient(135deg, rgba(23, 162, 184, 0.15) 0%, rgba(23, 162, 184, 0.1) 100%);
        border-left-color: #22b8cf;
        color: #4dabf7;
    }
}

.section-divider {
    border: none;
    height: 2px;
    background: linear-gradient(to right, var(--primary-color), transparent);
    margin: 2rem 0;
    opacity: 0.5;
}

/* Smooth transition untuk dark mode */
* {
    transition: background-color 0.3s ease, color 0.3s ease, border-color 0.3s ease;
}

/* Mobile responsive */
@media (max-width: 768px) {
    .hero-section h1 {
        font-size: 1.8rem;
    }
    
    .hero-section p {
        font-size: 0.95rem;
    }
    
    .stats-section {
        flex-direction: column;
    }
    
    .stat-box {
        min-width: 100%;
    }
}
</style>
""", unsafe_allow_html=True)

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
    
    **✨ Fitur Baru:** Aplikasi sekarang mendukung **Dark Mode** untuk kenyamanan mata Anda! 
    Dark mode akan otomatis menyesuaikan dengan preferensi sistem Anda.
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
    st.metric("🌙 Dark Mode", "✅")

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
        
        # Render setiap tool sebagai card
        st.markdown(f"""
        <div class="tool-card">
            <div class="tool-title">{desc.split()[0]} {' '.join(desc.split()[1:])}</div>
            <div class="tool-description">{desc}</div>
        </div>
        """, unsafe_allow_html=True)

# Info Section
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

col1, col2 = st.columns([2, 1])
with col1:
    st.markdown("""
    <div class="info-banner">
    <strong>💡 Tips Penggunaan</strong><br>
    • Gunakan sidebar untuk navigasi antar tool<br>
    • Setiap tool memiliki bantuan interaktif, hover/klik untuk melihat detail<br>
    • Selalu backup data sebelum melakukan proses besar<br>
    • Untuk hasil optimal, ikuti instruksi sesuai format yang diminta<br>
    • 🌙 Aktifkan Dark Mode melalui system settings untuk pengalaman optimal
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="info-banner info-banner-info">
    <strong>📞 Dukungan</strong><br>
    <small>Ada pertanyaan atau bug?<br>
    📧 nazarudin@gsid.co.id<br>
    Respon cepat & helpful support</small>
    </div>
    """, unsafe_allow_html=True)

footer("PGD Apps • Made by Nazarudin Zaini :D • Dark Mode Enabled ✨")