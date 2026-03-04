import streamlit as st

from utils import set_page, header, footer
from utils.auth import require_role
from utils.database import list_users

require_role("admin")
set_page("PGD Apps — Admin Panel", "🛡️")
header("🛡️ Admin Panel")

st.write("Halaman ini hanya bisa diakses oleh role **admin**.")
st.dataframe(list_users(), use_container_width=True)

footer("Admin")
