"""
app/models/user.py
-------------------
The User model backing authentication.

Inherits from UserMixin so Flask-Login gets is_authenticated,
is_active, is_anonymous, and get_id() for free.

Passwords are never stored in plain text — only a salted hash
(Werkzeug's generate_password_hash, which uses PBKDF2-SHA256 by
default) is persisted.
"""

from datetime import datetime, timezone
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

from app.database.db import db


class User(UserMixin, db.Model):
    __tablename__ = "users"

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False, index=True)
    email = db.Column(db.String(120), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(255), nullable=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    # --- Password helpers ---------------------------------------------
    def set_password(self, raw_password: str) -> None:
        """Hash and store the given plain-text password."""
        self.password_hash = generate_password_hash(raw_password)

    def check_password(self, raw_password: str) -> bool:
        """Verify a plain-text password against the stored hash."""
        return check_password_hash(self.password_hash, raw_password)

    def __repr__(self):
        return f"<User {self.username}>"
