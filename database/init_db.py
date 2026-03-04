from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from utils.database import create_user_with_password, get_user_by_username, init_db, list_users


def bootstrap_admin(username: str, password: str) -> None:
    init_db()
    if get_user_by_username(username):
        print(f"User '{username}' sudah ada, lewati pembuatan.")
        return
    create_user_with_password(username=username, password=password, role="admin")
    print(f"Admin '{username}' berhasil dibuat.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Initialize SQLite DB and seed first admin")
    parser.add_argument("--username", default="admin", help="Admin username")
    parser.add_argument("--password", default="admin123", help="Admin password")
    args = parser.parse_args()

    bootstrap_admin(args.username, args.password)
    print("Daftar user saat ini:")
    for user in list_users():
        print(user)
