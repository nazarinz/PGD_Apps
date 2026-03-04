import streamlit as st
import pandas as pd

from utils import set_page, header, footer
from utils.auth import require_role
from utils.database import (
    list_users,
    update_user_role,
    toggle_user_active,
    reset_password,
    delete_user,
    create_user_with_password
)

# =========================
# PAGE SETUP
# =========================

require_role("admin")

set_page("PGD Apps — Admin Panel", "🛡️")
header("🛡️ Admin Panel")

st.write("Kelola user aplikasi.")

# =========================
# CREATE USER
# =========================

st.subheader("➕ Create User")

col1, col2, col3 = st.columns(3)

with col1:
    new_username = st.text_input("Username")

with col2:
    new_password = st.text_input("Password", type="password")

with col3:
    new_role = st.selectbox("Role", ["user", "admin"])

if st.button("Create User"):

    if not new_username or not new_password:
        st.warning("Username dan password wajib diisi")

    else:
        try:
            create_user_with_password(
                new_username,
                new_password,
                new_role
            )

            st.success("User berhasil dibuat")
            st.rerun()

        except Exception:
            st.error("Username sudah digunakan")

# =========================
# LOAD USERS
# =========================

users = list_users()
df = pd.DataFrame(users)

if df.empty:
    st.info("Belum ada user.")
    st.stop()

df["is_active"] = df["is_active"].astype(bool)

# =========================
# USER TABLE
# =========================

st.subheader("👥 User Management")

edited_df = st.data_editor(
    df,
    use_container_width=True,
    num_rows="fixed",
    column_config={
        "id": st.column_config.NumberColumn("ID", disabled=True),
        "username": st.column_config.TextColumn("Username", disabled=True),
        "role": st.column_config.SelectboxColumn(
            "Role",
            options=["admin", "user"],
        ),
        "is_active": st.column_config.CheckboxColumn("Active"),
        "created_at": st.column_config.TextColumn("Created", disabled=True),
    },
)

# =========================
# SAVE CHANGES
# =========================

if st.button("💾 Save Changes"):

    for index, row in edited_df.iterrows():

        original = df.loc[index]

        if row["role"] != original["role"]:
            update_user_role(row["id"], row["role"])

        if row["is_active"] != original["is_active"]:
            toggle_user_active(row["id"], int(row["is_active"]))

    st.success("Perubahan disimpan")
    st.rerun()

# =========================
# RESET PASSWORD
# =========================

st.subheader("🔑 Reset Password")

user_options = {u["username"]: u["id"] for u in users}

selected_user = st.selectbox(
    "Pilih user",
    options=list(user_options.keys())
)

new_password = st.text_input(
    "Password baru",
    type="password",
    key="reset_pass"
)

if st.button("Reset Password"):

    if not new_password:
        st.warning("Password tidak boleh kosong")

    else:
        reset_password(user_options[selected_user], new_password)
        st.success("Password berhasil direset")

# =========================
# DELETE USER
# =========================

st.subheader("🗑 Delete User")

delete_user_name = st.selectbox(
    "Pilih user yang akan dihapus",
    options=list(user_options.keys()),
    key="delete_user"
)

if st.button("Delete User"):

    if delete_user_name == "admin":
        st.error("Admin utama tidak boleh dihapus")

    else:
        delete_user(user_options[delete_user_name])
        st.warning("User dihapus")
        st.rerun()

footer("Admin")