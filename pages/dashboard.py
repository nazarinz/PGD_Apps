import streamlit as st

from utils import set_page, header, footer
from utils.auth import require_login

require_login()
set_page("PGD Apps — Dashboard", "📊")
header("📊 User Dashboard")

st.write("Selamat datang di dashboard user. Halaman ini bisa diakses semua user yang sudah login.")

footer("Dashboard")
