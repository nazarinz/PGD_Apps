import streamlit as st
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo  # Python 3.9+

def set_page(title: str, icon: str = "🤖"):
    # Try to use custom PGD logo if available
    logo_path = Path(__file__).resolve().parent.parent / "assets" / "logo.png"
    page_icon = str(logo_path) if logo_path.exists() else icon
    try:
        st.set_page_config(page_title=title, page_icon=page_icon, layout="wide")
    except Exception:
        # set_page_config can only be called once; ignore if already set
        pass

def header(title: str, subtitle: str | None = None):
    st.title(title)
    if subtitle:
        st.caption(subtitle)

def footer(note: str = "PGD Apps • Made by Nazarudin Zaini :D"):
    st.markdown("---")
    left, right = st.columns([1,1])
    with left:
        st.caption(note)
    with right:
        # Waktu WIB (Jakarta) lengkap dengan detik
        now_wib = datetime.now(ZoneInfo("Asia/Jakarta"))
        st.caption(now_wib.strftime("%Y-%m-%d"))
