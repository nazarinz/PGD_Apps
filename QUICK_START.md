# 🚀 Quick Start Guide

Panduan cepat untuk memulai menggunakan PGD Apps.

## ⚡ 5 Menit Setup

### 1️⃣ Install Python (jika belum)
- Download dari https://www.python.org/downloads/
- Pilih Python 3.9 atau lebih tinggi
- **Penting:** Centang "Add Python to PATH" saat instalasi

### 2️⃣ Buka PowerShell
```powershell
# Navigasi ke folder aplikasi
cd "d:\Nazarudin Zaini\Dev Website Streamlit\PGD_Apps-main\PGD_Apps-main"
```

### 3️⃣ Buat Virtual Environment
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

### 4️⃣ Install Dependencies
```powershell
pip install -r requirements.txt
```

### 5️⃣ Jalankan Aplikasi
```powershell
streamlit run Home.py
```

✅ **Selesai!** Aplikasi akan terbuka di browser Anda.

---

## 🎯 Navigasi Dasar

### Home Page (Halaman Pertama)
- Tampilan overview semua tools
- Klik pada tool card untuk pergi ke halaman tersebut
- Lihat statistics dashboard

### Sidebar Navigation
- Klik icon ☰ (hamburger menu) jika sidebar tersembunyi
- Pilih tool dari dropdown "Pages"
- Setiap tool memiliki halaman tersendiri

### Footer
- Email: nazarudin@gsid.co.id untuk pertanyaan
- Tanggal dan waktu WIB

---

## 📚 Tools Overview

### 🧾 1. Quantity Change Extractor
**Gunakan untuk:** Ekstrak data quantity dari berbagai format
- Upload file (text, HTML, Excel)
- Data akan dinormalisasi otomatis
- Export hasil ke Excel

### 📋 2. Input PGD WFM BTP Tracking Report
**Gunakan untuk:** Tracking report harian
- Input data tracking
- Validasi otomatis
- Export report

### 📦 3. Merger Daily Report
**Gunakan untuk:** Gabung multiple file
- Upload multiple Excel files
- Hasil digabung dalam satu file
- Download hasil merger

### 📅 4. Jadwal Audit
**Gunakan untuk:** Generate jadwal audit
- Pilih periode (mingguan/bulanan)
- Sistem akan generate jadwal otomatis
- Export ke Excel

### 🔄 5. Reroute Tools
**Gunakan untuk:** Analisis PO changes
- Upload PO lama dan baru
- Sistem bandingkan dan analisis
- Lihat perbedaan dan konsistensi

### ⏳ 6. Input Tracking Report Pending Cancel
**Gunakan untuk:** Tracking status pending
- Input data pending/cancel
- Monitor progress
- Export status report

### 📏 7. Susun Sizelist
**Gunakan untuk:** Standardisasi daftar size
- Input size list
- Validasi dan standardisasi
- Export hasil

### 🔧 8. Tooling Sizelist
**Gunakan untuk:** Kelola tooling sizes
- Manage tooling size data
- Validasi data
- Export dan update

### 📧 9. Rekap E-Memo
**Gunakan untuk:** Buat recap email/memo
- Input email/memo content
- Sistem extract informasi penting
- Generate recap report

### 🌍 10. Check Export Plan Daily and Monthly
**Gunakan untuk:** Analisis SO export
- Check SO daily vs monthly
- Identifikasi discrepancy
- Generate report

### 🔍 11. Comparison RSA
**Gunakan untuk:** Bandingkan data RSA
- Upload data untuk dibandingkan
- Analisis perbedaan
- Visual comparison

### 💾 12 & 13. SAP INF DB Merger
**Gunakan untuk:** Merger database SAP
- Merger database files
- Validasi data
- Backup otomatis

---

## 💡 Tips & Tricks

### Eksport Data
Semua tools mendukung export Excel:
1. Setelah proses selesai, cari tombol "Download"
2. File akan terunduh otomatis
3. Format Excel sudah siap pakai (header, filter, freeze)

### Upload File
- Format yang didukung: XLS, XLSX, CSV, TXT
- Size maksimal: 200 MB
- Pastikan file sudah sesuai format yang diminta

### Error/Problem
- Baca pesan error yang ditampilkan dengan cermat
- Cek format file Anda
- Hubungi: nazarudin@gsid.co.id untuk bantuan

### Performance
- Jangan upload file terlalu besar sekaligus
- Tutup tab lain jika proses lambat
- Buka browser baru jika perlu restart

---

## 🎨 UI/UX Features

### Modern Design
- Color scheme yang profesional (blue primary)
- Card-based layout
- Hover effects yang smooth
- Responsive di semua ukuran layar

### Reusable Components
Jika Anda developer, gunakan komponen reusable:
```python
from utils import render_card, render_alert, render_progress_bar
```

Lihat `UI_UX_GUIDE.md` untuk dokumentasi lengkap.

---

## 🆘 Troubleshooting

### Aplikasi tidak mau start
```powershell
# Pastikan virtual environment aktif
.\venv\Scripts\Activate.ps1

# Coba install ulang dependencies
pip install -r requirements.txt --upgrade

# Run dengan debug mode
streamlit run Home.py --logger.level=debug
```

### Port 8501 sudah terpakai
```powershell
streamlit run Home.py --server.port=8502
```

### File tidak bisa diupload
- Cek ukuran file (max 200 MB)
- Cek format file (XLS, XLSX, CSV, TXT)
- Buka browser lain dan coba ulang

### Data tidak sesuai harapan
- Cek format data di Excel Anda
- Baca instruksi di halaman tool
- Lihat example data jika tersedia

---

## 📞 Need Help?

📧 **Email:** nazarudin@gsid.co.id
- Response time: Biasanya < 24 jam
- Jelaskan masalah dengan detail
- Sertakan screenshot jika perlu

---

## 📚 Further Reading

- `README.md` — Dokumentasi lengkap
- `UI_UX_GUIDE.md` — Panduan UI/UX components
- `EXAMPLE_COMPONENTS.py` — Contoh implementasi
- `CHANGELOG.md` — Histori perubahan

---

**Happy using PGD Apps! 🚀**

*Last Updated: December 10, 2025*
