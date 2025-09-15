# Home.py
import streamlit as st
from utils import set_page, header, footer
from pathlib import Path

set_page("PGD Apps — Home", "🧰")

header("🧰 PGD Apps — Home")
col1, col2 = st.columns([1,3])
with col1:
    logo = Path(__file__).resolve().parent / "assets" / "logo.png"
    if logo.exists():
        st.image(str(logo), width=140)
with col2:
    st.markdown("Selamat datang di **PGD Apps** — kumpulan tools harian tim PGD.")

st.subheader("📚 Daftar Halaman")
pages_dir = Path(__file__).resolve().parent / "pages"
if pages_dir.exists():
    items = sorted(pages_dir.glob("*.py"))
    descriptions = {
        "1_Quantity Change Tools": "Ekstrak/normalisasi quantity, reshape UK_* dan bandingkan perubahan qty.",
        "2_Check Export Plan Daily and Monthly": "Check SO rencana export harian vs bulanan, dengan kasus yang ada.",
        "3_Merger Daily Report": "Rekap banyak file.",
        "4_Jadwal Audit": "Generator jadwal audit mingguan/bulanan dengan format siap pakai.",
        "5_Reroute Tools": "Bandingkan Old vs New PO, cek konsistensi size, dan PO Finder batch.",
        "6_Input Tracking Report": "Tracking input/report status pekerjaan dan ekspor hasilnya.",
        "7_Susun Sizelist": "Susun dan standarkan daftar size (sizelist) sesuai kebutuhan produksi.",
    }
    for p in items:
        name = p.stem
        desc = descriptions.get(name, "")
        st.markdown(f"- **{name}** — {desc}")

st.info("Jika ada fitur yang ingin diubah atau ditambahkan email aja nazarudin@gsid.co.id, tetap semangat ygy.")
footer("PGD Apps • Made by Nazarudin Zaini :D")