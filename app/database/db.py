"""
db.py
-----
Holds the single SQLAlchemy instance used across the whole app.

Defining `db` here (rather than inside app/__init__.py) avoids circular
imports: models can do `from app.database.db import db` without needing
to import the app factory itself.
"""

from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()
