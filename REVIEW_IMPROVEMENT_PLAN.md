# Review: Area Improvement untuk PGD Apps

Dokumen ini merangkum area yang paling layak ditingkatkan berdasarkan pembacaan struktur repo, dokumentasi, dan pola implementasi saat ini.

## Prioritas Tinggi (Quick Win)

1. **Rapikan artefak repo (cache, bytecode, local env files)**
   - Tambahkan `.gitignore` standar Python agar file build/cache tidak ikut ter-commit.
   - Hapus file `__pycache__` yang sudah terlanjur masuk repo.
   - Dampak: repo lebih bersih, review PR lebih fokus ke source code.

2. **Sinkronisasi dokumentasi vs realita project**
   - README menyebut “13 tools”, tetapi folder `pages/` saat ini berisi lebih banyak halaman.
   - Beberapa nama file di bagian “Struktur Folder” juga tidak sama persis dengan kondisi aktual.
   - Dampak: onboarding user/dev lebih cepat, mengurangi miskomunikasi.

3. **Standarisasi quality gate sederhana (lint + test minimal)**
   - Tambahkan baseline check seperti:
     - `python -m compileall .`
     - optional: `ruff check .`
   - Jalankan lewat CI sederhana (GitHub Actions) saat push/PR.
   - Dampak: error sintaks atau typo cepat tertangkap sebelum production.

## Prioritas Menengah

4. **Refactor pola berulang di halaman Streamlit (`pages/`)**
   - Pola upload file, validasi kolom, normalisasi data, dan export Excel cenderung berulang.
   - Buat helper layer yang lebih konsisten di `utils/` agar page file lebih tipis.
   - Dampak: maintenance lebih murah dan bug fix bisa sekali-perbaiki untuk banyak page.

5. **Perkuat error handling yang user-facing**
   - Standardisasi pesan error/warning berbasis komponen UI reusable.
   - Pastikan setiap proses I/O (upload/parse/export) punya fallback message yang jelas.
   - Dampak: user lebih paham aksi yang harus dilakukan saat terjadi error.

6. **Tambahkan telemetry sederhana non-intrusif**
   - Contoh: log page usage count, waktu proses, dan jumlah baris input.
   - Simpan agregat lokal/internal untuk evaluasi performa tool paling sering dipakai.
   - Dampak: roadmap improvement lebih data-driven.

## Prioritas Strategis

7. **Test data contract per tool**
   - Buat test untuk memastikan schema input utama setiap tool tervalidasi.
   - Fokus awal di tool yang paling kritikal (misal rekap/merger/comparison).
   - Dampak: perubahan format data dari sumber upstream lebih cepat terdeteksi.

8. **Versi API internal untuk utilitas utama**
   - Tetapkan antarmuka stabil untuk helper (mis. formatter, mapper, exporter).
   - Catat breaking/non-breaking change di changelog.
   - Dampak: scaling tim lebih aman saat banyak kontributor edit utilitas yang sama.

9. **Perbaiki struktur dokumentasi agar berbasis use-case**
   - Pisahkan dokumen untuk:
     - user operasional (cara pakai)
     - developer (arsitektur, style guide)
     - maintainer (release checklist)
   - Dampak: navigasi knowledge base lebih efektif.

---

## Rekomendasi eksekusi 2 minggu

- **Minggu 1**: hygiene repo + sinkronisasi README + baseline checks.
- **Minggu 2**: refactor helper bersama untuk 2–3 halaman paling sering digunakan + test contract dasar.

Dengan urutan ini, tim bisa dapat peningkatan kualitas yang terasa tanpa mengganggu delivery harian.
