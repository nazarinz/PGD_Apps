# 🤖 PGD Apps — Multi-Page Streamlit Application

Kumpulan tools otomasi profesional untuk tim PGD dengan antarmuka modern dan responsif.

![Python](https://img.shields.io/badge/Python-3.9%2B-blue)
![Streamlit](https://img.shields.io/badge/Streamlit-1.37%2B-red)
![License](https://img.shields.io/badge/License-Private-green)

---

## 📋 Daftar Isi
- [Fitur Utama](#fitur-utama)
- [Instalasi](#instalasi)
- [Cara Menjalankan](#cara-menjalankan)
- [Tools & Fungsi](#tools--fungsi)
- [Struktur Folder](#struktur-folder)
- [Teknologi yang Digunakan](#teknologi-yang-digunakan)
- [UI/UX Improvements](#uiux-improvements)
- [Support & Maintenance](#support--maintenance)

---

## ✨ Fitur Utama

✅ **User-Friendly Interface** — Design yang modern dan intuitif
✅ **Fast Performance** — Proses data dengan cepat dan efisien
✅ **Multi-Tool Integration** — 13 tools berbeda dalam satu aplikasi
✅ **Excel Export** — Export otomatis dengan formatting yang rapi
✅ **Data Validation** — Validasi data otomatis sebelum proses
✅ **Responsive Design** — Bekerja sempurna di desktop, tablet, dan mobile
✅ **Dark Mode Ready** — Siap untuk dark theme Streamlit

---

## 🚀 Instalasi

### Prerequisites
- Python 3.9 atau lebih tinggi
- pip (Python package manager)

### Step-by-Step Installation

#### 1. Clone atau Download Repository
```bash
cd "d:\Nazarudin Zaini\Dev Website Streamlit\PGD_Apps-main\PGD_Apps-main"
```

#### 2. Buat Virtual Environment (Recommended)
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

#### 3. Install Dependencies
```powershell
pip install -r requirements.txt
```

---

## 🔐 Konfigurasi Auth untuk Deploy

Set minimal secrets/env berikut di platform deploy (mis. **Streamlit Cloud → Secrets**):

```toml
ADMIN_USERNAME = "admin"
ADMIN_PASSWORD = "ganti_dengan_password_kuat"
auth_secret = "random_panjang_32_atau_lebih"
```

Catatan:
- `auth_secret` dibaca dari `st.secrets.get("auth_secret")` atau env `AUTH_SECRET`.
- Jika `AUTH_SECRET` tidak ada, aplikasi membuat secret acak runtime per proses (lebih aman dibanding secret statis).

---

## ▶️ Cara Menjalankan

### Windows (PowerShell)
```powershell
cd "d:\Nazarudin Zaini\Dev Website Streamlit\PGD_Apps-main\PGD_Apps-main"
streamlit run Home.py
```

### Linux/Mac
```bash
cd /path/to/PGD_Apps
streamlit run Home.py
```

Aplikasi akan otomatis membuka di browser default Anda di `http://localhost:8501`

---

## 🛠️ Tools & Fungsi

| # | Tool | Fungsi | Icon |
|---|------|--------|------|
| 1 | **Quantity Change Extractor** | Ekstrak & normalisasi quantity dari berbagai format, reshape UK_*, bandingkan perubahan | 🧾 |
| 2 | **PGD WFM BTP Tracking** | Input & kelola tracking report dengan sistem terstruktur | 📋 |
| 3 | **Merger Daily Report** | Rekap & merger multiple file menjadi satu output | 📦 |
| 4 | **Jadwal Audit** | Generate jadwal audit mingguan/bulanan siap pakai | 📅 |
| 5 | **Reroute Tools** | Bandingkan PO, cek size consistency, PO Finder batch | 🔄 |
| 6 | **Input Tracking Report** | Tracking status pending/cancel & export hasil | ⏳ |
| 7 | **Susun Sizelist** | Standardisasi daftar size sesuai kebutuhan | 📏 |
| 8 | **Tooling Sizelist** | Kelola sizelist tooling dengan validasi | 🔧 |
| 9 | **Rekap E-Memo** | Rekap email/memo jadi laporan terintegrasi | 📧 |
| 10 | **Check Export Plan** | Analisis SO export harian vs bulanan | 🌍 |
| 11 | **Comparison RSA** | Analisis & bandingkan data RSA | 🔍 |
| 12 | **SAP INF DB Merger** | Merger database SAP dengan validasi | 💾 |
| 13 | **SAP INF DB Merger v2** | Versi lanjutan dengan fitur advanced | 💾 |

---

## 📁 Struktur Folder

```
PGD_Apps/
├── Home.py                          # Halaman utama aplikasi
├── README.md                        # File ini
├── UI_UX_GUIDE.md                   # Dokumentasi UI/UX components
├── requirements.txt                 # Dependencies list
├── assets/
│   ├── logo.png                     # Logo aplikasi
│   └── README.txt                   # Asset notes
├── pages/
│   ├── 1_Quantity Change Extractor.py
│   ├── 2_Input PGD_WFM_BTP_Tracking_Report.py
│   ├── 3_Merger Daily Report.py
│   ├── 4_Jadwal Audit.py
│   ├── 5_Reroute Tools.py
│   ├── 6_Input Tracking Report Pending Cancel.py
│   ├── 7_Susun Sizelist.py
│   ├── 8_Tooling Sizelist.py
│   ├── 9_Rekap E-Memo.py
│   ├── 10_Check Export Plan Daily and Monthly.py
│   ├── 11_Comparison RSA.py
│   ├── 12_SAP INF DB Merger.py
│   └── 13_SAP INF DB Mergerr.py
└── utils/
    ├── __init__.py                  # Package initialization
    ├── ui.py                        # Styling & layout functions
    ├── excel.py                     # Excel export utilities
    ├── components.py                # Reusable UI components
    ├── theme.py                     # Theme configuration
    └── __pycache__/                 # Cache folder (auto-generated)
```

---

## 🔧 Teknologi yang Digunakan

| Library | Versi | Fungsi |
|---------|-------|--------|
| Streamlit | 1.37+ | Web framework |
| Pandas | 2.1+ | Data manipulation |
| NumPy | 1.26+ | Numerical computing |
| Openpyxl | 3.1+ | Excel reading/writing |
| XlsxWriter | 3.2+ | Excel formatting |
| BeautifulSoup4 | 4.12+ | HTML parsing |
| LXML | 5.2+ | XML/HTML processing |
| Python-dateutil | 2.9+ | Date utilities |
| Holidays | 0.60+ | Holiday calendar |

---

## 🎨 UI/UX Improvements (v1.1)

Dokumentasi lengkap tersedia di [`UI_UX_GUIDE.md`](UI_UX_GUIDE.md)

### Apa yang Baru:

✅ **Enhanced Styling** — Custom CSS untuk semua elemen
✅ **Modern Color Scheme** — Blue primary color dengan good contrast
✅ **Reusable Components** — 9+ komponen UI yang dapat digunakan kembali
✅ **Better Home Page** — Hero section, card-based layout, statistics
✅ **Improved Typography** — Font Segoe UI dengan sizing yang optimal
✅ **Hover Effects** — Smooth transitions & animations
✅ **Theme Configuration** — Centralized color & spacing management
✅ **Responsive Design** — Mobile-first approach
✅ **Better Excel Export** — Formatted header, frozen panes, autofilter

### Component List:
- `render_card()` — Card container
- `render_stats()` — Statistics grid
- `render_section_divider()` — Section divider
- `render_progress_bar()` — Progress indicator
- `render_alert()` — Alert messages
- `render_tabs()` — Tabbed interface
- `render_help_box()` — Help/tip box
- `render_code_block()` — Code display
- `render_data_quality_indicator()` — Data quality visual

---

## 📝 Requirements.txt

Semua dependensi sudah terdaftar di `requirements.txt`:

```
streamlit>=1.37
pandas>=2.1
numpy>=1.26
openpyxl>=3.1
xlsxwriter>=3.2
xlrd>=2.0.1
pyxlsb>=1.0.10
odfpy>=1.4.1
beautifulsoup4>=4.12
lxml>=5.2
python-dateutil>=2.9
holidays>=0.60
```

---

## 🐛 Troubleshooting

### Aplikasi tidak buka di browser
```powershell
streamlit run Home.py --logger.level=debug
```

### Error: ModuleNotFoundError
```powershell
pip install -r requirements.txt --upgrade
```

### Port 8501 sudah terpakai
```powershell
streamlit run Home.py --server.port=8502
```

---

## 🤝 Contributing & Support

### Lapor Bug atau Request Fitur
📧 Email: **nazarudin@gsid.co.id**

### Development Guidelines
1. Maintain backward compatibility
2. Gunakan reusable components dari `utils/`
3. Follow existing code style
4. Update documentation jika ada perubahan

---

## 📄 Lisensi

Private - Tim PGD Only

---

## 👨‍💻 Author

**Nazarudin Zaini**
- Email: nazarudin@gsid.co.id
- Role: PGD Apps Developer

---

## 📅 Version History

| Versi | Tanggal | Changes |
|-------|---------|---------|
| 1.0 | 2024 | Initial release |
| 1.1 | Dec 10, 2025 | UI/UX improvements, new components |

---

**Happy Coding! 🚀**