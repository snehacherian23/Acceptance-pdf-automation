"""
app/routes/auth_routes.py
--------------------------
Blueprint handling registration, login, and logout.

Kept deliberately thin: validation lives in the form classes
(app/utils/forms.py), password handling lives on the User model
itself. This route file's job is just to orchestrate the flow and
talk to Flask-Login.
"""

from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_user, logout_user, login_required, current_user

from app.database.db import db
from app.models.user import User
from app.utils.forms import RegisterForm, LoginForm

auth_bp = Blueprint("auth", __name__, url_prefix="/auth")


@auth_bp.route("/register", methods=["GET", "POST"])
def register():
    # Already logged in? Don't let them register again.
    if current_user.is_authenticated:
        return redirect(url_for("dashboard.index"))

    form = RegisterForm()

    if form.validate_on_submit():
        new_user = User(username=form.username.data.strip(), email=form.email.data.strip().lower())
        new_user.set_password(form.password.data)

        db.session.add(new_user)
        db.session.commit()

        flash("Account created successfully. You can now log in.", "success")
        return redirect(url_for("auth.login"))

    return render_template("auth/register.html", form=form)


@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard.index"))

    form = LoginForm()

    if form.validate_on_submit():
        user = User.query.filter_by(username=form.username.data.strip()).first()

        if user is not None and user.check_password(form.password.data):
            login_user(user, remember=form.remember_me.data)
            flash(f"Welcome back, {user.username}!", "success")

            # Support "next" redirect (e.g. hit a protected page while logged out)
            next_page = request.args.get("next")
            return redirect(next_page or url_for("dashboard.index"))

        flash("Invalid username or password.", "danger")

    return render_template("auth/login.html", form=form)


@auth_bp.route("/logout")
@login_required
def logout():
    logout_user()
    flash("You have been logged out.", "info")
    return redirect(url_for("auth.login"))
