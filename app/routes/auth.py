import time

import bcrypt
from flask import Blueprint, render_template, request, redirect, url_for, session, flash

from ..db import get_bh_db

auth_bp = Blueprint("auth", __name__)

# ── Login rate limiting (per client IP, in-process) ───────────────────────────
LOGIN_MAX_ATTEMPTS = 10
LOGIN_WINDOW_SECS = 300

_login_attempts: dict = {}


def _rate_limit_ok(ip: str) -> bool:
    now = time.time()
    attempts = [t for t in _login_attempts.get(ip, []) if now - t < LOGIN_WINDOW_SECS]
    if len(attempts) >= LOGIN_MAX_ATTEMPTS:
        _login_attempts[ip] = attempts
        return False
    attempts.append(now)
    _login_attempts[ip] = attempts
    return True


@auth_bp.route("/", methods=["GET", "POST"])
@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    if "user" in session:
        return redirect(url_for("dashboard.overview"))
    if request.method == "POST":
        ip = request.remote_addr or "0.0.0.0"
        if not _rate_limit_ok(ip):
            flash("Too many login attempts. Please try again in a few minutes.")
            return render_template("login.html"), 429

        username = request.form.get("username", "").strip()[:64]
        password = request.form.get("password", "").strip()[:256]
        db = get_bh_db()
        user = db.dashboard_users.find_one({"username": username})
        stored = (user or {}).get("password", "")
        if user and stored and bcrypt.checkpw(password.encode(), stored.encode()):
            # Drop any pre-login session state before granting access.
            session.clear()
            session.permanent = True
            session["user"] = username
            session["role"] = user.get("role", "admin")
            session["stores"] = user.get("stores", [])
            _login_attempts.pop(ip, None)
            return redirect(url_for("dashboard.overview"))
        flash("Invalid credentials")
    return render_template("login.html")


@auth_bp.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("auth.login"))
