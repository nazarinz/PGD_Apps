# Home.py
import streamlit as st
from utils import set_page, header, footer
from pathlib import Path

set_page("PGD Apps — Home", "🤖")

# Custom CSS untuk home page yang lebih menarik
st.markdown("""
<style>
.hero-section {
    background: linear-gradient(135deg, #1f77b4 0%, #0d47a1 100%);
    color: white;
    padding: 3rem 2rem;
    border-radius: 10px;
    margin-bottom: 2rem;
    text-align: center;
}

.hero-section h1 {
    color: white;
    border: none;
    margin: 0;
    font-size: 2.5rem;
}

.hero-section p {
    font-size: 1.1rem;
    margin: 0.5rem 0 0 0;
    opacity: 0.95;
}

.tool-card {
    background: linear-gradient(135deg, #f8f9fa 0%, #ffffff 100%);
    border: 1px solid #e0e0e0;
    border-radius: 8px;
    padding: 1.5rem;
    margin: 1rem 0;
    transition: all 0.3s ease;
    box-shadow: 0 2px 6px rgba(0, 0, 0, 0.08);
}

.tool-card:hover {
    transform: translateY(-4px);
    box-shadow: 0 8px 20px rgba(31, 119, 180, 0.15);
    border-color: #1f77b4;
}

.tool-icon {
    font-size: 1.8rem;
    margin-right: 0.5rem;
}

.tool-title {
    color: #0d47a1;
    font-size: 1.1rem;
    font-weight: 600;
    margin-bottom: 0.5rem;
}

.tool-description {
    color: #555;
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
}

.stat-number {
    font-size: 2rem;
    color: #1f77b4;
    font-weight: bold;
}

.stat-label {
    color: #666;
    font-size: 0.9rem;
    margin-top: 0.5rem;
}

.info-banner {
    background: linear-gradient(135deg, #fff3cd 0%, #fffbea 100%);
    border-left: 4px solid #ffc107;
    padding: 1rem;
    border-radius: 6px;
    margin: 2rem 0;
}

.section-divider {
    border: none;
    height: 2px;
    background: linear-gradient(to right, #1f77b4, transparent);
    margin: 2rem 0;
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
    """)

# Stats Section
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("📊 Tools Tersedia", 7)
with col2:
    st.metric("⚡ Fitur Utama", 15)
with col3:
    st.metric("📈 Efisiensi", "40%+")
with col4:
    st.metric("👥 Tim", "1 Dev")

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
        "12_SAP INF DB Merger": "💾 Merger database SAP INF dengan sistem validasi dan backup otomatis.",
        "13_SAP INF DB Mergerr": "💾 Merger database SAP INF versi lanjutan dengan fitur advanced.",
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
    • Untuk hasil optimal, ikuti instruksi sesuai format yang diminta
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="info-banner" style="border-left-color: #17a2b8; background: linear-gradient(135deg, #d1ecf1 0%, #e8f4f8 100%);">
    <strong>📞 Dukungan</strong><br>
    <small>Ada pertanyaan atau bug?<br>
    📧 nazarudin@gsid.co.id<br>
    Respon cepat & helpful support</small>
    </div>
    """, unsafe_allow_html=True)

footer("PGD Apps • Made by Nazarudin Zaini :D")