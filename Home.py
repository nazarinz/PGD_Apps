"""Home page for PGD Apps — v1.1.0
Layout modern dengan system theme support (Light/Dark otomatis).
Versi ini memperbaiki: stat tool count, tool card rendering, deskripsi lengkap 13 tools,
logo fallback, placeholder text, 2-kolom tool grid, dan version badge.
"""
import streamlit as st
from utils import set_page, header, footer
from utils.auth import init_auth_state, render_sidebar_auth
from pathlib import Path

set_page("PGD Apps — Home", "🤖")
init_auth_state()
render_sidebar_auth()

# ============================================================
# JavaScript: sync background dengan mode theme pilihan user (System/Light/Dark)
# ============================================================
st.markdown("""
<script>
    const applyTheme = () => {
        const userTheme = document.documentElement.getAttribute('data-user-theme');
        const isDark = userTheme
            ? userTheme === 'dark'
            : window.matchMedia('(prefers-color-scheme: dark)').matches;
        const root = document.documentElement;
        const stApp = document.querySelector('.stApp');
        const main = document.querySelector('.main');
        const bg = isDark ? '#0f172a' : '#ffffff';
        root.style.setProperty('background-color', bg, 'important');
        if (stApp) stApp.style.backgroundColor = bg;
        if (main) main.style.backgroundColor = bg;
    };
    setTimeout(applyTheme, 100);
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', applyTheme);
</script>
""", unsafe_allow_html=True)

# ============================================================
# CSS: Light/Dark variables + komponen styling
# ============================================================
css = """
<style>
/* ── Root variables (Light Mode) ── */
:root {
    --primary-color:        #2563eb;
    --primary-dark:         #1d4ed8;
    --accent-color:         #0ea5e9;
    --text-primary:         #0f172a;
    --text-secondary:       #475569;
    --text-muted:           #94a3b8;
    --bg-primary:           #ffffff;
    --bg-secondary:         #f8fafc;
    --card-bg:              #fafbff;
    --border-color:         #e0e7ff;
    --shadow:               rgba(37, 99, 235, 0.08);
    --hero-start:           #3b82f6;
    --hero-end:             #2563eb;
    color-scheme: light;
}

/* ── Dark Mode override ── */
@media (prefers-color-scheme: dark) {
    :root {
        --primary-color:    #60a5fa;
        --primary-dark:     #3b82f6;
        --accent-color:     #93c5fd;
        --text-primary:     #f1f5f9;
        --text-secondary:   #cbd5e1;
        --text-muted:       #64748b;
        --bg-primary:       #0f172a;
        --bg-secondary:     #1e293b;
        --card-bg:          #1e293b;
        --border-color:     #334155;
        --shadow:           rgba(96, 165, 250, 0.15);
        --hero-start:       #1e40af;
        --hero-end:         #3730a3;
        color-scheme: dark;
    }
}

:root[data-user-theme='light'] {
    --primary-color:        #2563eb;
    --primary-dark:         #1d4ed8;
    --accent-color:         #0ea5e9;
    --text-primary:         #0f172a;
    --text-secondary:       #475569;
    --text-muted:           #94a3b8;
    --bg-primary:           #ffffff;
    --bg-secondary:         #f8fafc;
    --card-bg:              #fafbff;
    --border-color:         #e0e7ff;
    --shadow:               rgba(37, 99, 235, 0.08);
    --hero-start:           #3b82f6;
    --hero-end:             #2563eb;
    color-scheme: light;
}

:root[data-user-theme='dark'] {
    --primary-color:    #60a5fa;
    --primary-dark:     #3b82f6;
    --accent-color:     #93c5fd;
    --text-primary:     #f1f5f9;
    --text-secondary:   #cbd5e1;
    --text-muted:       #64748b;
    --bg-primary:       #0f172a;
    --bg-secondary:     #1e293b;
    --card-bg:          #1e293b;
    --border-color:     #334155;
    --shadow:           rgba(96, 165, 250, 0.15);
    --hero-start:       #1e40af;
    --hero-end:         #3730a3;
    color-scheme: dark;
}

/* ── Global backgrounds ── */
html, body,
.stApp,
section[data-testid="stAppViewContainer"],
section[data-testid="stAppViewContainer"] > .main,
.main .block-container {
    background-color: var(--bg-primary) !important;
    background:       var(--bg-primary) !important;
    color: var(--text-primary) !important;
}

/* ── Sidebar ── */
[data-testid="stSidebar"],
[data-testid="stSidebar"] > div:first-child {
    background-color: var(--bg-secondary) !important;
    border-right: 1px solid var(--border-color);
}
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
    color: var(--text-primary) !important;
}

/* ── Hero Section ── */
.hero-section {
    background: linear-gradient(135deg, var(--hero-start) 0%, var(--hero-end) 100%);
    color: white;
    padding: 3.5rem 2rem 3rem;
    border-radius: 16px;
    margin-bottom: 2.5rem;
    text-align: center;
    box-shadow: 0 8px 32px var(--shadow);
    border: 1px solid rgba(255,255,255,0.1);
    position: relative;
    overflow: hidden;
}
.hero-section::after {
    content: '';
    position: absolute;
    top: -50%;
    right: -10%;
    width: 300px;
    height: 300px;
    background: rgba(255,255,255,0.05);
    border-radius: 50%;
    pointer-events: none;
}
.hero-section h1 {
    font-size: 2.5rem;
    font-weight: 700;
    margin-bottom: 0.75rem;
    text-shadow: 0 2px 4px rgba(0,0,0,0.15);
    color: white !important;
}
.hero-section p {
    font-size: 1.15rem;
    opacity: 0.92;
    color: white !important;
    margin-bottom: 0;
}
.hero-badge {
    display: inline-block;
    background: rgba(255,255,255,0.18);
    color: white !important;
    border: 1px solid rgba(255,255,255,0.3);
    border-radius: 20px;
    padding: 0.25rem 0.85rem;
    font-size: 0.8rem;
    font-weight: 600;
    letter-spacing: 0.05em;
    margin-bottom: 1rem;
    backdrop-filter: blur(4px);
}

/* ── About section ── */
.about-box {
    background: var(--card-bg);
    border: 1.5px solid var(--border-color);
    border-radius: 12px;
    padding: 1.5rem 1.75rem;
    height: 100%;
    box-shadow: 0 2px 8px var(--shadow);
}
.about-box h3 {
    color: var(--primary-color) !important;
    margin-bottom: 0.75rem;
}
.about-box p, .about-box li {
    color: var(--text-secondary) !important;
    line-height: 1.7;
    font-size: 0.95rem;
}

/* ── Logo placeholder ── */
.logo-placeholder {
    display: flex;
    align-items: center;
    justify-content: center;
    width: 140px;
    height: 140px;
    background: linear-gradient(135deg, var(--hero-start), var(--hero-end));
    border-radius: 20px;
    font-size: 3.5rem;
    box-shadow: 0 4px 16px var(--shadow);
    margin: auto;
}

/* ── Stats ── */
.stat-box {
    text-align: center;
    padding: 1.5rem 1rem;
    background: var(--card-bg);
    border-radius: 12px;
    border: 1.5px solid var(--border-color);
    box-shadow: 0 1px 3px rgba(37,99,235,0.06), 0 4px 8px rgba(37,99,235,0.04);
    transition: transform 0.2s ease, box-shadow 0.2s ease;
    cursor: default;
}
.stat-box:hover {
    transform: translateY(-3px);
    box-shadow: 0 6px 16px rgba(37,99,235,0.12), 0 10px 24px rgba(37,99,235,0.08);
}
.stat-number {
    font-size: 2.2rem;
    color: var(--primary-color);
    font-weight: 800;
    margin-bottom: 0.25rem;
    line-height: 1;
}
.stat-label {
    color: var(--text-secondary);
    font-size: 0.85rem;
    font-weight: 500;
    margin-top: 0.35rem;
}

/* ── Tool Cards ── */
.tool-card {
    background: var(--card-bg);
    border: 1.5px solid var(--border-color);
    border-radius: 12px;
    padding: 1.25rem 1.5rem;
    margin-bottom: 1rem;
    transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease;
    box-shadow: 0 1px 3px rgba(37,99,235,0.06);
    position: relative;
    overflow: hidden;
}
.tool-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0;
    width: 4px; height: 100%;
    background: var(--primary-color);
    opacity: 0;
    transition: opacity 0.2s ease;
}
.tool-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 6px 18px rgba(37,99,235,0.12), 0 12px 32px rgba(37,99,235,0.07);
    border-color: var(--primary-color);
}
.tool-card:hover::before { opacity: 1; }

.tool-header {
    display: flex;
    align-items: flex-start;
    gap: 0.75rem;
    margin-bottom: 0.4rem;
}
.tool-number {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    min-width: 26px;
    height: 26px;
    background: var(--primary-color);
    color: white !important;
    border-radius: 6px;
    font-size: 0.7rem;
    font-weight: 700;
    flex-shrink: 0;
    margin-top: 2px;
}
.tool-title {
    color: var(--primary-color) !important;
    font-weight: 700;
    font-size: 1rem;
    line-height: 1.4;
    margin: 0;
}
.tool-description {
    color: var(--text-secondary) !important;
    line-height: 1.6;
    font-size: 0.9rem;
    margin: 0;
    padding-left: 2.1rem;
}

/* ── Info Banner ── */
.info-banner {
    background: var(--card-bg);
    border-left: 4px solid var(--primary-color);
    padding: 1.25rem 1.5rem;
    border-radius: 8px;
    margin: 0;
    box-shadow: 0 2px 8px var(--shadow);
}
.info-banner.accent { border-left-color: var(--accent-color); }
.info-banner strong { color: var(--primary-color) !important; font-size: 1rem; }
.info-banner p, .info-banner li, .info-banner small {
    color: var(--text-secondary) !important;
    font-size: 0.9rem;
    line-height: 1.7;
}

/* ── Divider ── */
.section-divider {
    border: none;
    height: 2px;
    background: linear-gradient(to right, var(--primary-color), transparent);
    margin: 2.25rem 0;
    opacity: 0.25;
}

/* ── Buttons ── */
.stButton > button {
    background: linear-gradient(135deg, var(--primary-color), var(--primary-dark));
    color: white !important;
    border-radius: 8px;
    font-weight: 600;
    border: none;
    transition: all 0.2s ease;
}
.stButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px var(--shadow);
}

/* ── Text color override ── */
.stMarkdown p, .stMarkdown li,
h1, h2, h3, h4, h5, h6 {
    color: var(--text-primary) !important;
}

/* ── DataFrames ── */
.stDataFrame, table.dataframe {
    background: var(--card-bg) !important;
    color: var(--text-primary) !important;
}

/* ── Smooth transitions ── */
* { transition: background-color 0.2s ease, color 0.2s ease, border-color 0.2s ease; }

/* ── Mobile ── */
@media (max-width: 768px) {
    .hero-section { padding: 2rem 1rem 1.75rem; }
    .hero-section h1 { font-size: 1.8rem; }
    .hero-section p  { font-size: 0.95rem; }
    .tool-description { padding-left: 0; }
}
</style>
"""
st.markdown(css, unsafe_allow_html=True)

header("🤖 PGD Apps — Home")

# ============================================================
# Hero Section
# ============================================================
st.markdown("""
<div class="hero-section">
    <div style="display:flex; justify-content:center; gap:0.5rem; flex-wrap:wrap; margin-bottom:1rem;">
        <div class="hero-badge">✨ Version 1.1.0 — December 2025</div>
        <div class="hero-badge" style="background:rgba(251,191,36,0.25); border-color:rgba(251,191,36,0.5);">⚠️ Unofficial — Bukan Produk Resmi</div>
    </div>
    <h1>🚀 Selamat Datang di PGD Apps</h1>
    <p>Kumpulan 13 tools otomasi harian untuk tim PGD — cepat, akurat, dan user-friendly</p>
</div>
""", unsafe_allow_html=True)

# ============================================================
# Unofficial Disclaimer Banner
# ============================================================
st.markdown("""
<div style="
    background: rgba(251,191,36,0.1);
    border: 1px solid rgba(251,191,36,0.4);
    border-left: 4px solid #f59e0b;
    border-radius: 8px;
    padding: 0.75rem 1.25rem;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: flex-start;
    gap: 0.6rem;
">
    <span style="font-size:1.1rem;">⚠️</span>
    <div>
        <strong style="color:#b45309 !important; font-size:0.9rem;">Aplikasi Tidak Resmi (Unofficial)</strong>
        <p style="margin:0.2rem 0 0; font-size:0.85rem; color:#92400e !important; line-height:1.5;">
            PGD Apps adalah tools pribadi yang dibuat untuk membantu efisiensi kerja tim,
            bukan merupakan produk resmi dari perusahaan. Penggunaan sepenuhnya menjadi
            tanggung jawab pengguna. Selalu verifikasi hasil dengan data sumber asli.
        </p>
    </div>
</div>
""", unsafe_allow_html=True)

# ============================================================
# About & Logo Section
# ============================================================
col_logo, col_about = st.columns([1, 3])

with col_logo:
    logo_path = Path(__file__).resolve().parent / "assets" / "logo.png"
    if logo_path.exists():
        st.image(str(logo_path), width=140, caption="PGD Logo")
    else:
        st.markdown("""
        <div class="logo-placeholder">🤖</div>
        """, unsafe_allow_html=True)

with col_about:
    st.markdown("""
    <div class="about-box">
        <h3>💼 Tentang PGD Apps</h3>
        <p>
            PGD Apps adalah platform tools otomasi terintegrasi yang dirancang khusus untuk
            meningkatkan efisiensi operasional tim PGD. Dengan antarmuka yang modern dan
            intuitif, setiap proses manual kini dapat dikerjakan lebih cepat dan akurat.
        </p>
        <p>
            <strong>✨ Catatan:</strong> Tampilan mengikuti preferensi sistem Anda (Light/Dark Mode).
            Untuk mengubah, cukup atur tema OS atau browser Anda.
        </p>
    </div>
    """, unsafe_allow_html=True)

# ============================================================
# Stats Section
# ============================================================
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

col1, col2, col3, col4 = st.columns(4)
stats = [
    ("13",   "📊 Tools Tersedia"),
    ("50+",  "⚡ Fitur Aktif"),
    ("40%",  "📈 Efisiensi Kerja"),
    ("Auto", "🌙 Tema System"),
]
for col, (num, label) in zip([col1, col2, col3, col4], stats):
    with col:
        st.markdown(f"""
        <div class="stat-box">
            <div class="stat-number">{num}</div>
            <div class="stat-label">{label}</div>
        </div>
        """, unsafe_allow_html=True)

# ============================================================
# Tools Section
# ============================================================
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
st.subheader("📚 Daftar Tools & Fitur")

# Deskripsi lengkap 13 tools — key = stem nama file di folder pages/
TOOL_INFO = {
    "1_Quantity Change Extractor": {
        "icon": "🧾",
        "title": "Quantity Change Extractor",
        "desc": "Ekstrak dan normalisasi data quantity dari berbagai format, reshape kolom UK_*, serta bandingkan perubahan qty secara akurat dan otomatis.",
    },
    "2_Input PGD_WFM_BTP_Tracking_Report": {
        "icon": "📋",
        "title": "Input PGD WFM BTP Tracking Report",
        "desc": "Input dan kelola data tracking report PGD WFM BTP dengan sistem terstruktur, validasi otomatis, dan export siap pakai.",
    },
    "3_Merger Daily Report": {
        "icon": "📦",
        "title": "Merger Daily Report",
        "desc": "Rekap dan merger banyak file laporan harian menjadi satu output Excel yang rapi dengan header dan format standar.",
    },
    "4_Jadwal Audit": {
        "icon": "📅",
        "title": "Jadwal Audit",
        "desc": "Generator jadwal audit mingguan dan bulanan dengan format siap pakai, mudah diexport dan langsung bisa digunakan.",
    },
    "5_Reroute Tools": {
        "icon": "🔄",
        "title": "Reroute Tools",
        "desc": "Bandingkan Old vs New PO, periksa konsistensi ukuran (size), serta batch PO Finder otomatis untuk efisiensi kerja.",
    },
    "6_Input Tracking Report Pending Cancel": {
        "icon": "⏳",
        "title": "Input Tracking Report Pending / Cancel",
        "desc": "Tracking status pekerjaan pending dan cancel secara terstruktur, dengan kemampuan ekspor hasil ke berbagai format.",
    },
    "7_Susun Sizelist": {
        "icon": "📏",
        "title": "Susun Sizelist",
        "desc": "Susun dan standarisasi daftar ukuran (sizelist) sesuai kebutuhan produksi dengan validasi otomatis dan format yang konsisten.",
    },
    "8_Tooling Sizelist": {
        "icon": "🔧",
        "title": "Tooling Sizelist",
        "desc": "Kelola sizelist tooling lengkap dengan fitur import, export, validasi data, dan deteksi inkonsistensi secara otomatis.",
    },
    "9_Rekap E-Memo": {
        "icon": "📧",
        "title": "Rekap E-Memo",
        "desc": "Rekap data dari email atau memo internal dan buat laporan terintegrasi yang siap digunakan dengan mudah dan cepat.",
    },
    "10_Check Export Plan Daily and Monthly": {
        "icon": "🌍",
        "title": "Check Export Plan Daily & Monthly",
        "desc": "Periksa dan bandingkan SO rencana export harian vs bulanan, serta identifikasi kasus dan inkonsistensi yang ada.",
    },
    "11_Comparison RSA": {
        "icon": "🔍",
        "title": "Comparison RSA",
        "desc": "Analisis dan bandingkan data RSA secara mendalam dengan statistik performa yang terperinci dan visualisasi hasil.",
    },
    "12_SAP INF DB Merger": {
        "icon": "💾",
        "title": "SAP INF DB Merger",
        "desc": "Merger dan konsolidasi database SAP dengan validasi integritas data, penanganan duplikat, dan backup otomatis.",
    },
    "13_SAP INF DB Mergerr": {
        "icon": "💾",
        "title": "SAP INF DB Merger (Advanced)",
        "desc": "Versi lanjutan SAP INF DB Merger dengan fitur pemrosesan lebih lengkap, analisis mendalam, dan opsi konfigurasi tambahan.",
    },
    # Fallback untuk Advanced Analytics jika masih ada
    "12_Advanced_Analytics": {
        "icon": "📊",
        "title": "Advanced Analytics",
        "desc": "Dashboard analitik lanjutan dengan visualisasi data interaktif, statistik mendalam, dan laporan komprehensif.",
    },
}

pages_dir = Path(__file__).resolve().parent / "pages"

if pages_dir.exists():
    # Sort NUMERIK berdasarkan angka di awal nama file (1, 2, 3 ... 10, 11, dst)
    def _sort_key(p):
        try:
            return int(p.stem.split("_")[0])
        except ValueError:
            return 9999

    items = sorted(pages_dir.glob("*.py"), key=_sort_key)

    # Render 2 kolom
    left_items = []
    right_items = []
    for i, p in enumerate(items):
        if i % 2 == 0:
            left_items.append(p)
        else:
            right_items.append(p)

    col_left, col_right = st.columns(2)

    def render_tool_card(p, col):
        name = p.stem
        info = TOOL_INFO.get(name)
        if info:
            icon  = info["icon"]
            title = info["title"]
            desc  = info["desc"]
        else:
            # Fallback: bersihkan nama file menjadi judul yang rapi
            icon  = "🛠️"
            parts = name.split("_")
            # Buang bagian angka di depan
            title_parts = parts[1:] if parts[0].isdigit() else parts
            title = " ".join(title_parts)
            desc  = "Klik tool ini di sidebar untuk melihat detail fitur dan cara penggunaannya."

        # Ambil nomor dari awal nama file
        num = name.split("_")[0] if name and name[0].isdigit() else "–"

        with col:
            st.markdown(f"""
            <div class="tool-card" role="article" aria-label="{title}">
                <div class="tool-header">
                    <span class="tool-number">{num}</span>
                    <p class="tool-title">{icon} {title}</p>
                </div>
                <p class="tool-description">{desc}</p>
            </div>
            """, unsafe_allow_html=True)

    for p in left_items:
        render_tool_card(p, col_left)

    for p in right_items:
        render_tool_card(p, col_right)

else:
    st.warning("⚠️ Folder `pages/` tidak ditemukan. Pastikan struktur direktori sudah benar.")

# ============================================================
# Tips & Support Section
# ============================================================
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

col_tips, col_support = st.columns([2, 1])

with col_tips:
    st.markdown("""
    <div class="info-banner" role="note">
        <strong>💡 Tips Penggunaan</strong>
        <p style="margin-top:0.5rem; margin-bottom:0;">
        • Gunakan <b>sidebar</b> untuk navigasi antar tool dengan cepat<br>
        • Setiap tool dilengkapi bantuan interaktif — perhatikan kotak info di tiap halaman<br>
        • Selalu <b>backup data</b> sebelum melakukan proses besar<br>
        • Untuk hasil optimal, ikuti format kolom yang diminta di tiap tool<br>
        • File Excel hasil export sudah siap pakai: header berwarna, filter, dan freeze panes
        </p>
    </div>
    """, unsafe_allow_html=True)

with col_support:
    st.markdown("""
    <div class="info-banner accent" role="note">
        <strong>📞 Dukungan</strong>
        <p style="margin-top:0.5rem; margin-bottom:0;">
        Ada pertanyaan, bug, atau request fitur?<br><br>
        📧 <b>nazarudin@gsid.co.id</b><br><br>
        <small>Respon biasanya &lt; 24 jam</small>
        </p>
    </div>
    """, unsafe_allow_html=True)

# ============================================================
# Footer
# ============================================================
footer("PGD Apps v1.1.0 • Made with ❤️ by Nazarudin Zaini")
