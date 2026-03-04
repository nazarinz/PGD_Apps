from __future__ import annotations

import os

import streamlit as st

from utils import database as db
from utils.auth import hash_password


def _get_setting(key: str, default: str | None = None) -> str | None:
    """Read setting with priority: Streamlit secrets -> environment variable -> default."""
    try:
        if key in st.secrets:
            return str(st.secrets.get(key))
    except Exception:
        pass
    return os.getenv(key, default)


def bootstrap_admin_if_empty() -> None:
    """Initialize DB and create first admin user when users table is empty."""
    db.init_db()

    if db.list_users():
        return

    admin_username = _get_setting("ADMIN_USERNAME", "admin")
    admin_password = _get_setting("ADMIN_PASSWORD")

    if not admin_password:
        st.error("ADMIN_PASSWORD belum diset di secrets/env. Set dulu supaya admin bisa dibuat otomatis.")
        st.stop()

    if db.get_user_by_username(admin_username):
        return

    create_user_with_password = getattr(db, "create_user_with_password", None)
    if callable(create_user_with_password):
        create_user_with_password(username=admin_username, password=admin_password, role="admin")
    else:
        legacy_create_user = getattr(db, "create_user", None)
        if not callable(legacy_create_user):
            st.error("Fungsi pembuatan user tidak ditemukan di utils.database.")
            st.stop()

        legacy_create_user(
            username=admin_username,
            password_hash=hash_password(admin_password),
            role="admin",
        )

    st.success(f"Bootstrap selesai: admin '{admin_username}' dibuat (first run). Silakan login.")
