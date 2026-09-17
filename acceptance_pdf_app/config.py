"""
config.py
---------
Central configuration for the Flask application.

Keeps environment-specific settings (secret key, database URI, upload
limits, folder paths) out of main.py and out of the app factory, so the
same codebase can run in development or production just by changing the
FLASK_ENV / environment variables.
"""

import os
from datetime import timedelta

# Absolute path to the project root (the folder this file lives in)
BASE_DIR = os.path.abspath(os.path.dirname(__file__))


class BaseConfig:
    """Shared configuration for every environment."""

    # --- Security ---------------------------------------------------
    # In production this MUST be overridden via an environment variable.
    SECRET_KEY = os.environ.get("SECRET_KEY", "dev-secret-key-change-me")

    # --- Database -----------------------------------------------------
    SQLALCHEMY_DATABASE_URI = os.environ.get(
        "DATABASE_URL", f"sqlite:///{os.path.join(BASE_DIR, 'instance', 'app.db')}"
    )
    SQLALCHEMY_TRACK_MODIFICATIONS = False

    # --- Session --------------------------------------------------
    PERMANENT_SESSION_LIFETIME = timedelta(hours=8)
    SESSION_COOKIE_HTTPONLY = True
    SESSION_COOKIE_SAMESITE = "Lax"

    # --- File Uploads -------------------------------------------------
    UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
    PROCESSED_FOLDER = os.path.join(BASE_DIR, "processed")

    # Batches can include several large Excel sheets; 50 MB is a safe ceiling.
    MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50 MB

    ALLOWED_EXCEL_EXTENSIONS = {"xlsx", "xls"}


class DevelopmentConfig(BaseConfig):
    DEBUG = True


class ProductionConfig(BaseConfig):
    DEBUG = False


class TestingConfig(BaseConfig):
    TESTING = True
    SQLALCHEMY_DATABASE_URI = "sqlite:///:memory:"


# Map string names (used via FLASK_ENV) to config classes
config_by_name = {
    "development": DevelopmentConfig,
    "production": ProductionConfig,
    "testing": TestingConfig,
}
