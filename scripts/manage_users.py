"""
Dashboard User Management Script
=================================
Add, list, or reset passwords for dashboard users.

Usage:
    python scripts/manage_users.py add       <username> <password>
    python scripts/manage_users.py reset     <username> <new_password>
    python scripts/manage_users.py list
    python scripts/manage_users.py delete    <username>
"""

import sys
import os

# Allow running from project root
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import bcrypt
from dotenv import load_dotenv
from pymongo import MongoClient

load_dotenv()

MONGO_URI = os.environ.get("MONGO_URI", "")


def get_db():
    if not MONGO_URI:
        raise SystemExit("MONGO_URI is not set — copy .env.example to .env and fill it in")
    client = MongoClient(MONGO_URI, serverSelectionTimeoutMS=5000)
    return client[os.environ.get("DB_NAME", "beinghumanServer")]


def add_user(username: str, password: str):
    db = get_db()
    if db.dashboard_users.find_one({"username": username}):
        print(f"[ERROR] User '{username}' already exists. Use 'reset' to change the password.")
        return
    hashed = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
    db.dashboard_users.insert_one({"username": username, "password": hashed, "role": "admin", "stores": []})
    print(f"[OK] User '{username}' created successfully.")


def reset_password(username: str, new_password: str):
    db = get_db()
    user = db.dashboard_users.find_one({"username": username})
    if not user:
        print(f"[ERROR] User '{username}' not found.")
        return
    hashed = bcrypt.hashpw(new_password.encode(), bcrypt.gensalt()).decode()
    db.dashboard_users.update_one({"username": username}, {"$set": {"password": hashed}})
    print(f"[OK] Password reset for '{username}'.")


def list_users():
    db = get_db()
    users = db.dashboard_users.find({}, {"username": 1, "_id": 0})
    user_list = list(users)
    if not user_list:
        print("[INFO] No dashboard users found.")
        return
    print(f"[INFO] {len(user_list)} dashboard user(s):")
    for u in user_list:
        print(f"  - {u['username']}")


def delete_user(username: str):
    db = get_db()
    result = db.dashboard_users.delete_one({"username": username})
    if result.deleted_count:
        print(f"[OK] User '{username}' deleted.")
    else:
        print(f"[ERROR] User '{username}' not found.")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    command = sys.argv[1].lower()

    if command == "add" and len(sys.argv) == 4:
        add_user(sys.argv[2], sys.argv[3])
    elif command == "reset" and len(sys.argv) == 4:
        reset_password(sys.argv[2], sys.argv[3])
    elif command == "list":
        list_users()
    elif command == "delete" and len(sys.argv) == 3:
        delete_user(sys.argv[2])
    else:
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
