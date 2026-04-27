"""
CIS School System - Authentication Routes
==========================================
Handles: login, logout, change password
"""

from flask import Blueprint, render_template, request, redirect, url_for, flash, session
from werkzeug.security import check_password_hash, generate_password_hash
from ..models import db, Teacher

auth_bp = Blueprint("auth", __name__)


def login_required(f):
    """Decorator — redirects to login if user is not logged in"""
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if "teacher_id" not in session:
            flash("Please log in to continue.", "warning")
            return redirect(url_for("auth.login"))
        return f(*args, **kwargs)
    return decorated


def role_required(*roles):
    """Decorator — restricts a route to specific roles"""
    def decorator(f):
        from functools import wraps
        @wraps(f)
        def decorated(*args, **kwargs):
            if "teacher_id" not in session:
                return redirect(url_for("auth.login"))
            teacher = Teacher.query.get(session["teacher_id"])
            if teacher is None or teacher.role not in roles:
                flash("You do not have permission to access that page.", "danger")
                return redirect(url_for("main.dashboard"))
            return f(*args, **kwargs)
        return decorated
    return decorator


@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    if "teacher_id" in session:
        return redirect(url_for("main.dashboard"))

    if request.method == "POST":
        email    = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")

        teacher = Teacher.query.filter_by(email=email, is_active=True).first()

        if teacher and check_password_hash(teacher.password_hash, password):
            session["teacher_id"]   = teacher.id
            session["teacher_name"] = teacher.full_name
            session["teacher_role"] = teacher.role
            flash(f"Welcome back, {teacher.full_name}!", "success")
            return redirect(url_for("main.dashboard"))
        else:
            flash("Invalid email or password. Please try again.", "danger")

    return render_template("auth/login.html")


@auth_bp.route("/logout")
def logout():
    session.clear()
    flash("You have been logged out.", "info")
    return redirect(url_for("auth.login"))


@auth_bp.route("/change-password", methods=["GET", "POST"])
@login_required
def change_password():
    if request.method == "POST":
        current  = request.form.get("current_password", "")
        new_pw   = request.form.get("new_password", "")
        confirm  = request.form.get("confirm_password", "")

        teacher = Teacher.query.get(session["teacher_id"])

        if not check_password_hash(teacher.password_hash, current):
            flash("Current password is incorrect.", "danger")
        elif new_pw != confirm:
            flash("New passwords do not match.", "danger")
        elif len(new_pw) < 6:
            flash("Password must be at least 6 characters.", "danger")
        else:
            teacher.password_hash = generate_password_hash(new_pw)
            db.session.commit()
            flash("Password changed successfully.", "success")
            return redirect(url_for("main.dashboard"))

    return render_template("auth/change_password.html")
