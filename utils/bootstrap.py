from __future__ import annotations

import os

import streamlit as st

from utils.database import create_user_with_password, get_user_by_username, init_db, list_users


def _get_setting(key: str, default: str | None = None) -> str | None:
    """Read setting with priority: Streamlit secrets -> environment variable -> default."""
    if key in st.secrets:
        return str(st.secrets.get(key))
    return os.getenv(key, default)


def bootstrap_admin_if_empty() -> None:
    """
    Run once at app startup:
    - ensure auth tables exist
    - if no user exists yet, create an admin user from secrets/env
    """
    init_db()

    if len(list_users()) > 0:
        return

    admin_username = _get_setting("ADMIN_USERNAME", "admin")
    admin_password = _get_setting("ADMIN_PASSWORD", None)

    if not admin_password:
        st.error("ADMIN_PASSWORD belum diset di secrets/env. Set dulu supaya admin bisa dibuat otomatis.")
        st.stop()

    if get_user_by_username(admin_username):
        return

    create_user_with_password(username=admin_username, password=admin_password, role="admin")
    st.success(f"Bootstrap selesai: admin '{admin_username}' dibuat (first run). Silakan login.")
