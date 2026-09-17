"""
app/__init__.py
----------------
The application factory.

Using the factory pattern (create_app) instead of a single global `app`
object means:
  - We can create multiple app instances with different configs
    (useful for testing).
  - Extensions (db, login_manager) are defined once and attached to
    whichever app instance is being built, avoiding circular imports
    between routes/services/models.
"""

import os
from flask import Flask
from flask_login import LoginManager

from config import config_by_name
from app.database.db import db

# --- Extensions (created here, initialized inside create_app) ---------
login_manager = LoginManager()
login_manager.login_view = "auth.login"          # redirect target for @login_required
login_manager.login_message = "Please log in to access the dashboard."
login_manager.login_message_category = "warning"


def create_app(config_name=None):
    """
    Application factory.

    Args:
        config_name: "development" | "production" | "testing".
                     Falls back to the FLASK_ENV env var, then "development".
    """
    if config_name is None:
        config_name = os.environ.get("FLASK_ENV", "development")

    app = Flask(__name__, instance_relative_config=True)
    app.config.from_object(config_by_name[config_name])

    # Make sure runtime folders exist (uploads/, processed/, instance/)
    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    os.makedirs(app.config["PROCESSED_FOLDER"], exist_ok=True)
    os.makedirs(app.instance_path, exist_ok=True)

    # --- Initialize extensions with this app instance -----------------
    db.init_app(app)
    login_manager.init_app(app)

    # --- User loader for Flask-Login -----------------------------------
    # Imported lazily inside the function body (not at module level) to
    # avoid a circular import: models/user.py will import `db` from this
    # package, so app/__init__.py must not import models/user.py at the
    # top of the file.
    @login_manager.user_loader
    def load_user(user_id):
        from app.models.user import User
        return User.query.get(int(user_id))

    # --- Register blueprints --------------------------------------
    # Blueprints are added incrementally as each module is built.
    # (auth_routes registered in Step 3, dashboard_routes in Step 4,
    #  processing_routes in Step 6+)
    _register_blueprints(app)

    # --- Create database tables on first run --------------------------
    with app.app_context():
        db.create_all()

    return app


def _register_blueprints(app):
    """Import and register blueprints. Wrapped in try/except per-blueprint
    so that earlier steps of the build (before a module exists yet) don't
    crash the whole app — each blueprint becomes active as soon as its
    file is created."""

    try:
        from app.routes.auth_routes import auth_bp
        app.register_blueprint(auth_bp)
    except ImportError:
        pass

    try:
        from app.routes.dashboard_routes import dashboard_bp
        app.register_blueprint(dashboard_bp)
    except ImportError:
        pass

    try:
        from app.routes.processing_routes import processing_bp
        app.register_blueprint(processing_bp)
    except ImportError:
        pass

    try:
        from app.routes.main_routes import main_bp
        app.register_blueprint(main_bp)
    except ImportError:
        pass
