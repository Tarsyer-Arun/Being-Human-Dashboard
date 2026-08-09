from flask import Blueprint, render_template, session, redirect, url_for
from functools import wraps

dashboard_bp = Blueprint("dashboard", __name__)

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user" not in session:
            return redirect(url_for("auth.login"))
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user" not in session:
            return redirect(url_for("auth.login"))
        if session.get("role", "admin") != "admin":
            return redirect(url_for("dashboard.overview"))
        return f(*args, **kwargs)
    return decorated

@dashboard_bp.route("/overview")
@login_required
def overview():
    return render_template("overview.html")

@dashboard_bp.route("/customer-unattended")
@login_required
def customer_unattended():
    return render_template("customer_unattended.html")
