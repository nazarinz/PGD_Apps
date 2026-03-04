from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any
import bcrypt

BASE_DIR = Path(__file__).resolve().parent.parent
DB_DIR = BASE_DIR / "database"
DB_PATH = DB_DIR / "users.db"
_DB_INITIALIZED = False


def _connect() -> sqlite3.Connection:
    """Create SQLite connection without running schema bootstrap."""
    DB_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def ensure_db_initialized() -> None:
    """Ensure auth tables are available before query execution."""
    global _DB_INITIALIZED
    if _DB_INITIALIZED:
        return

    init_db()
    _DB_INITIALIZED = True


def get_connection() -> sqlite3.Connection:
    """Create a SQLite connection with row access by column name."""
    ensure_db_initialized()
    return _connect()


def init_db() -> None:
    """Initialize required auth tables if not exists."""
    with _connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL CHECK(role IN ('admin', 'user')),
                is_active INTEGER NOT NULL DEFAULT 1,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        conn.commit()


# =========================
# USER QUERIES
# =========================

def get_user_by_username(username: str) -> sqlite3.Row | None:
    with get_connection() as conn:
        return conn.execute(
            "SELECT id, username, password_hash, role, is_active FROM users WHERE username = ?",
            (username,),
        ).fetchone()


def get_user_by_id(user_id: int) -> sqlite3.Row | None:
    with get_connection() as conn:
        return conn.execute(
            "SELECT id, username, role, is_active FROM users WHERE id = ?",
            (user_id,),
        ).fetchone()


def list_users() -> list[dict[str, Any]]:
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT id, username, role, is_active, created_at FROM users ORDER BY id"
        ).fetchall()

    return [dict(row) for row in rows]


# =========================
# CREATE USER
# =========================

def create_user_with_password(
    username: str,
    password: str,
    role: str = "user"
) -> int:

    password_hash = bcrypt.hashpw(
        password.encode(),
        bcrypt.gensalt()
    ).decode()

    with get_connection() as conn:
        cursor = conn.execute(
            "INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)",
            (username, password_hash, role),
        )

        conn.commit()
        return int(cursor.lastrowid)


# =========================
# USER MANAGEMENT
# =========================

def update_user_role(user_id: int, role: str) -> None:
    with get_connection() as conn:
        conn.execute(
            "UPDATE users SET role=? WHERE id=?",
            (role, user_id)
        )
        conn.commit()


def toggle_user_active(user_id: int, active: int) -> None:
    with get_connection() as conn:
        conn.execute(
            "UPDATE users SET is_active=? WHERE id=?",
            (active, user_id)
        )
        conn.commit()


def reset_password(user_id: int, new_password: str) -> None:

    password_hash = bcrypt.hashpw(
        new_password.encode(),
        bcrypt.gensalt()
    ).decode()

    with get_connection() as conn:
        conn.execute(
            "UPDATE users SET password_hash=? WHERE id=?",
            (password_hash, user_id)
        )
        conn.commit()


def delete_user(user_id: int) -> None:
    with get_connection() as conn:
        conn.execute(
            "DELETE FROM users WHERE id=?",
            (user_id,)
        )
        conn.commit()
