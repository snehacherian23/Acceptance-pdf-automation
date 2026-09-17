"""
app/routes/dashboard_routes.py
-------------------------------
Serves the dashboard page: upload form + results/log panel, wired to
the /processing/run pipeline via app/static/js/dashboard.js.

Functional-first version -- stat cards, progress bar animation, and
visual polish (icons, hover effects, etc.) come in a later pass, per
current priority of "get the pipeline working end-to-end."
"""

from flask import Blueprint, render_template
from flask_login import login_required

dashboard_bp = Blueprint("dashboard", __name__, url_prefix="/dashboard")


@dashboard_bp.route("/")
@login_required
def index():
    return render_template("dashboard/index.html")
