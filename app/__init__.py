import json
import os
from datetime import timedelta

from dotenv import load_dotenv
from flask import Flask, request, session

# Must run before .db is imported so that module reads the real environment.
load_dotenv()

from .db import init_db  # noqa: E402


def create_app():
    app = Flask(__name__)

    secret_key = os.environ.get("FLASK_SECRET_KEY", "")
    if not secret_key:
        raise RuntimeError(
            "FLASK_SECRET_KEY is not set. Copy .env.example to .env and fill in the values."
        )
    app.secret_key = secret_key

    # Session cookie hardening. Secure is on unless we're running locally over
    # plain HTTP, otherwise the browser would drop the cookie entirely.
    app.config["SESSION_COOKIE_HTTPONLY"] = True
    app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
    app.config["SESSION_COOKIE_SECURE"] = os.environ.get("FLASK_ENV") == "production"
    app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(
        hours=int(os.environ.get("SESSION_HOURS", 8))
    )
    app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024

    init_db(app)

    @app.after_request
    def set_security_headers(response):
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "SAMEORIGIN"
        response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
        if request.path.startswith("/api/"):
            response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate"
            response.headers["Pragma"] = "no-cache"
        return response

    @app.context_processor
    def inject_stores():
        stores_file = os.path.join(app.root_path, '..', 'stores.json')
        try:
            with open(stores_file, 'r') as f:
                all_stores = json.load(f)
        except Exception:
            all_stores = []

        role = session.get("role", "admin")
        user_stores = session.get("stores", [])

        if role != "user" or not user_stores:
            stores = all_stores
        else:
            allowed = set(user_stores)
            stores = [s for s in all_stores if s["code"] in allowed]

        return dict(stores=stores, session_role=role)

    from .routes.auth import auth_bp
    from .routes.dashboard import dashboard_bp
    from .routes.api import api_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(dashboard_bp)
    app.register_blueprint(api_bp, url_prefix="/api")

    return app
