from __future__ import annotations

import hashlib
import hmac
import os
from typing import Any

import bcrypt
import streamlit as st

from utils.database import get_user_by_id, get_user_by_username

SESSION_AUTH_KEY = "auth_user"
SESSION_AUTH_SIG = "auth_sig"
_RUNTIME_FALLBACK_SECRET = os.urandom(32).hex()


def _session_secret() -> str:
    return str(st.secrets.get("auth_secret") or os.getenv("AUTH_SECRET") or _RUNTIME_FALLBACK_SECRET)


def _build_signature(user_id: int, username: str, role: str) -> str:
    payload = f"{user_id}:{username}:{role}".encode("utf-8")
    return hmac.new(_session_secret().encode("utf-8"), payload, hashlib.sha256).hexdigest()


def init_auth_state() -> None:
    st.session_state.setdefault(SESSION_AUTH_KEY, None)
    st.session_state.setdefault(SESSION_AUTH_SIG, None)


def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")


def verify_password(password: str, password_hash: str) -> bool:
    return bcrypt.checkpw(password.encode("utf-8"), password_hash.encode("utf-8"))


def login(username: str, password: str) -> bool:
    init_auth_state()

    user = get_user_by_username(username)
    if not user:
        return False

    if int(user["is_active"]) != 1:
        return False

    if not verify_password(password, user["password_hash"]):
        return False

    session_user = {
        "user_id": int(user["id"]),
        "username": str(user["username"]),
        "role": str(user["role"]),
    }
    st.session_state[SESSION_AUTH_KEY] = session_user
    st.session_state[SESSION_AUTH_SIG] = _build_signature(
        session_user["user_id"], session_user["username"], session_user["role"]
    )
    return True


def _session_valid() -> bool:
    session_user = st.session_state.get(SESSION_AUTH_KEY)
    session_sig = st.session_state.get(SESSION_AUTH_SIG)

    if not isinstance(session_user, dict) or not isinstance(session_sig, str):
        return False

    required_keys = {"user_id", "username", "role"}
    if set(session_user.keys()) != required_keys:
        return False

    expected_sig = _build_signature(
        int(session_user["user_id"]),
        str(session_user["username"]),
        str(session_user["role"]),
    )
    if not hmac.compare_digest(session_sig, expected_sig):
        return False

    db_user = get_user_by_id(int(session_user["user_id"]))
    if not db_user:
        return False

    if int(db_user["is_active"]) != 1:
        return False

    if db_user["username"] != session_user["username"]:
        return False

    if db_user["role"] != session_user["role"]:
        return False

    return True


def is_logged_in() -> bool:
    init_auth_state()
    if not _session_valid():
        logout(redirect=False)
        return False
    return True


def get_current_user() -> dict[str, Any] | None:
    if not is_logged_in():
        return None
    return st.session_state.get(SESSION_AUTH_KEY)


def logout(redirect: bool = True) -> None:
    for key in (SESSION_AUTH_KEY, SESSION_AUTH_SIG):
        st.session_state.pop(key, None)

    if redirect:
        st.switch_page("Home.py")


def _render_login_form(form_key: str = "login_form") -> bool:
    with st.form(form_key, clear_on_submit=False):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submit = st.form_submit_button("Masuk", use_container_width=True)

    if not submit:
        return False

    if login(username.strip(), password):
        st.session_state["show_login_popup"] = False
        st.success("Login berhasil")
        st.rerun()

    st.error("Username atau password salah / akun nonaktif")
    return False


def open_login_popup(dialog_title: str = "🔐 Login PGD Apps", form_key: str = "login_form") -> None:
    @st.dialog(dialog_title)
    def _login_popup() -> None:
        _render_login_form(form_key=form_key)

    _login_popup()


def render_sidebar_auth() -> None:
    init_auth_state()
    st.session_state.setdefault("show_login_popup", False)
    st.session_state.setdefault("login_popup_auto_opened", False)

    # Buka popup otomatis sekali per sesi saat user belum login.
    if not is_logged_in() and not st.session_state["login_popup_auto_opened"]:
        st.session_state["show_login_popup"] = True
        st.session_state["login_popup_auto_opened"] = True

    left, right = st.columns([4, 1])
    with right:
        if is_logged_in():
            user = get_current_user() or {}
            st.caption(f"👤 **{user.get('username', '-')}** ({user.get('role', '-')})")
            if st.button("Logout", use_container_width=True, key="auth_logout_button"):
                st.session_state["show_login_popup"] = False
                st.session_state["login_popup_auto_opened"] = False
                logout()
            return

        if st.button("🔐 Login", use_container_width=True, key="auth_login_button"):
            st.session_state["show_login_popup"] = True

    if st.session_state.get("show_login_popup", False):
        open_login_popup(form_key="login_popup_form")


def require_login() -> None:
    if not is_logged_in():
        st.warning("Silakan login terlebih dahulu dari tombol 🔐 Login di halaman Home.")
        st.stop()


def require_role(role: str) -> None:
    require_login()
    user = get_current_user()
    if not user:
        st.stop()

    if user.get("role") != role:
        st.error("Anda tidak memiliki akses ke halaman ini.")
        st.stop()
