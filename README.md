# PGD Apps (Multi-Page Streamlit)
Ini adalah kumpulan utilitas PGD dalam bentuk aplikasi Streamlit multi-halaman.

Cara menjalankan (Windows/PowerShell):

1. Masuk ke folder app:
   cd PGD_Apps
2. (Opsional) Buat virtualenv dan aktifkan.
3. Install dependensi:
   pip install -r requirements.txt
4. Jalankan aplikasi (halaman Home):
   streamlit run Home.py

Halaman yang tersedia di sidebar:
- 1_PO_Tools.py — Extractor & Normalizer
- 2_SO_AutoDetect_Matcher.py — Pencocokan SO otomatis
- 3_HTML_XLS_Merger.py — Gabung beberapa HTML/XLS
- 4_Jadwal_Piket.py — Generator jadwal piket mingguan