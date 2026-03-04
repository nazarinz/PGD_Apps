from __future__ import annotations

import os

import streamlit as st

from utils import database as db


def _get_setting(key: str, default: str | None = None) -> str | None:
    """Read setting with priority: Streamlit secrets -> environment variable -> default."""
    if key in st.secrets:
        return str(st.secrets.get(key))
    return os.getenv(key, default)


def _create_admin_user(username: str, password: str) -> None:
    """Create admin user using available database API (new or legacy)."""
    create_user_with_password = getattr(db, "create_user_with_password", None)
    if callable(create_user_with_password):
        create_user_with_password(username=username, password=password, role="admin")
        return

    legacy_create_user = getattr(db, "create_user", None)
    if callable(legacy_create_user):
        from utils.auth import hash_password

        legacy_create_user(
            username=username,
            password_hash=hash_password(password),
            role="admin",
        )
        return

    st.error("Fungsi pembuatan user tidak ditemukan di utils.database.")
    st.stop()


def bootstrap_admin_if_empty() -> None:
    """
    Run once at app startup:
    - ensure auth tables exist
    - if no user exists yet, create an admin user from secrets/env
    """
    db.init_db()

    if len(db.list_users()) > 0:
        return

    admin_username = _get_setting("ADMIN_USERNAME", "admin")
    admin_password = _get_setting("ADMIN_PASSWORD", None)

    if not admin_password:
        st.error("ADMIN_PASSWORD belum diset di secrets/env. Set dulu supaya admin bisa dibuat otomatis.")
        st.stop()

    if db.get_user_by_username(admin_username):
        return

    _create_admin_user(username=admin_username, password=admin_password)
    st.success(f"Bootstrap selesai: admin '{admin_username}' dibuat (first run). Silakan login.")
