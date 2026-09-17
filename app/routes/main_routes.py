"""
app/routes/main_routes.py
----------------------------
Handles the site root ("/"). Without this, visiting the bare domain
(e.g. http://127.0.0.1:5000) 404s even though /dashboard and
/auth/login work fine -- this route is what makes the app feel like
a complete site rather than a collection of separate pages.
"""

from flask import Blueprint, redirect, url_for
from flask_login import current_user

main_bp = Blueprint("main", __name__)


@main_bp.route("/")
def root():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard.index"))
    return redirect(url_for("auth.login"))
