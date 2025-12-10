"""
CONTOH IMPLEMENTASI UI/UX COMPONENTS

File ini menunjukkan bagaimana menggunakan komponen UI baru di halaman Anda.
Copy dan modifikasi sesuai kebutuhan.
"""

import streamlit as st
from utils import (
    set_page, header, footer,
    render_card, render_stats, render_section_divider,
    render_progress_bar, render_alert, render_help_box,
    display_success_message, display_error_message,
    render_data_quality_indicator
)
import pandas as pd

# =====================================================
# Setup Page
# =====================================================
set_page("Example Page", "📝")
header("📝 Contoh Implementasi UI/UX Components")

# =====================================================
# Section 1: Help Box
# =====================================================
render_help_box(
    title="Panduan Menggunakan Halaman Ini",
    content="Halaman ini menunjukkan cara menggunakan semua komponen UI yang tersedia. "
            "Scroll ke bawah untuk melihat berbagai contoh implementasi."
)

render_section_divider()

# =====================================================
# Section 2: Alert Messages
# =====================================================
st.subheader("📢 Alert Messages")

col1, col2, col3, col4 = st.columns(4)

with col1:
    render_alert("Proses berhasil!", "success")

with col2:
    render_alert("Ada peringatan penting", "warning")

with col3:
    render_alert("Terjadi kesalahan", "error")

with col4:
    render_alert("Informasi penting", "info")

# Alternative: gunakan helper functions
st.markdown("### Menggunakan Helper Functions:")
col1, col2 = st.columns(2)
with col1:
    display_success_message("Data berhasil diupload!")
with col2:
    display_error_message("File format tidak valid!")

render_section_divider()

# =====================================================
# Section 3: Cards
# =====================================================
st.subheader("🎴 Card Components")

render_card(
    title="Data Processing Tool",
    content="Alat ini memproses data Excel dengan cepat dan akurat. "
            "Mendukung berbagai format file dan validasi otomatis.",
    footer="📊 Compatible dengan: XLS, XLSX, CSV",
    icon="⚙️"
)

render_card(
    title="Export Results",
    content="Hasil proses dapat diexport dalam berbagai format: "
            "Excel, CSV, atau JSON dengan formatting yang rapi.",
    footer="💾 Storage: Up to 200 MB per file",
    icon="💿"
)

render_card(
    title="Real-time Monitoring",
    content="Monitor progress processing secara real-time dengan visual indicators. "
            "Mendapat notifikasi ketika proses selesai.",
    footer="⏱️ Average processing time: 2-5 detik",
    icon="📈"
)

render_section_divider()

# =====================================================
# Section 4: Statistics
# =====================================================
st.subheader("📊 Statistics Cards")

stats = {
    "📁 Total Files": "1,245",
    "✅ Successfully Processed": "1,180",
    "❌ Failed": "65",
    "⏳ In Progress": "0"
}
render_stats(stats, cols=4)

render_section_divider()

# =====================================================
# Section 5: Progress Bars
# =====================================================
st.subheader("📈 Progress Indicators")

st.markdown("**File Upload Progress:**")
render_progress_bar(0.85, "Upload", "primary")

st.markdown("**Data Processing:**")
render_progress_bar(0.45, "Processing", "info")

st.markdown("**Export:**")
render_progress_bar(0.99, "Export", "success")

render_section_divider()

# =====================================================
# Section 6: Data Quality
# =====================================================
st.subheader("📋 Data Quality Indicators")

col1, col2, col3 = st.columns(3)

with col1:
    render_data_quality_indicator(92, "Data Completeness")

with col2:
    render_data_quality_indicator(75, "Data Accuracy")

with col3:
    render_data_quality_indicator(60, "Data Consistency")

render_section_divider()

# =====================================================
# Section 7: Sample Data
# =====================================================
st.subheader("📊 Sample Data Table")

df = pd.DataFrame({
    'ID': range(1, 6),
    'Name': ['Item A', 'Item B', 'Item C', 'Item D', 'Item E'],
    'Quantity': [100, 250, 175, 300, 150],
    'Status': ['✅ Done', '✅ Done', '⏳ Processing', '✅ Done', '❌ Failed'],
    'Percentage': [100, 100, 45, 100, 0]
})

st.dataframe(df, use_container_width=True)

# =====================================================
# Section 8: File Upload Example
# =====================================================
st.subheader("📤 File Upload Example")

uploaded_file = st.file_uploader(
    "Upload file Excel Anda:",
    type=['xlsx', 'xls', 'csv']
)

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        display_success_message(f"File '{uploaded_file.name}' berhasil dibaca!")
        st.markdown(f"**Total Rows:** {len(df)}")
        st.dataframe(df.head(10), use_container_width=True)
        
    except Exception as e:
        display_error_message(f"Error: {str(e)}")

render_section_divider()

# =====================================================
# Section 9: Form Example
# =====================================================
st.subheader("📝 Form Example dengan Styling")

with st.form("example_form"):
    col1, col2 = st.columns(2)
    
    with col1:
        name = st.text_input("Nama Lengkap:", placeholder="Masukkan nama Anda")
        email = st.text_input("Email:", placeholder="email@example.com")
    
    with col2:
        department = st.selectbox(
            "Departemen:",
            ["PGD", "Warehouse", "Finance", "HR"]
        )
        quantity = st.number_input("Quantity:", min_value=1, max_value=1000, value=10)
    
    notes = st.text_area("Catatan:", placeholder="Tambahkan catatan di sini...", height=100)
    
    submitted = st.form_submit_button("✅ Submit Form", use_container_width=True)
    
    if submitted:
        if name and email:
            display_success_message(f"Terima kasih {name}! Form Anda telah dikirim.")
        else:
            display_error_message("Silahkan isi semua field yang diperlukan!")

render_section_divider()

# =====================================================
# Footer
# =====================================================
footer()

# =====================================================
# CODE REFERENCE
# =====================================================
st.markdown("""
<hr style="margin-top: 3rem;">

### 📚 Referensi Kode

Untuk melihat dokumentasi lengkap semua komponen, lihat:
- **`UI_UX_GUIDE.md`** — Panduan lengkap komponen
- **`utils/components.py`** — Source code komponen
- **`utils/ui.py`** — Styling dan konfigurasi

### 🚀 Quick Start

```python
from utils import render_card, render_alert, display_success_message

# Render card
render_card("Title", "Content", "Footer", "🎯")

# Render alert
render_alert("Message", "success")

# Display message
display_success_message("Success!")
```

---
**Made with ❤️ by Nazarudin Zaini**
""", unsafe_allow_html=True)
