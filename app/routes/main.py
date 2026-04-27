"""
CIS School System - Main / Dashboard Routes
=============================================
The home page after login — shows a summary relevant to the logged-in role.
"""

from flask import Blueprint, render_template, session, redirect, url_for
from ..models import db, Teacher, Student, Grade, Term, Assessment, ReportCard
from .auth import login_required

main_bp = Blueprint("main", __name__)


@main_bp.route("/")
def index():
    if "teacher_id" in session:
        return redirect(url_for("main.dashboard"))
    return redirect(url_for("auth.login"))


@main_bp.route("/dashboard")
@login_required
def dashboard():
    teacher = Teacher.query.get(session["teacher_id"])
    role    = teacher.role

    # Active term
    active_term = Term.query.filter_by(is_active=True).first()
    assessments = []
    if active_term:
        assessments = Assessment.query.filter_by(term_id=active_term.id).order_by(Assessment.number).all()

    # Stats for admin / principal
    stats = {}
    if role in ("admin", "principal"):
        stats["total_students"]   = Student.query.filter_by(is_active=True).count()
        stats["total_teachers"]   = Teacher.query.filter_by(is_active=True).count()
        stats["grades_count"]     = Grade.query.count()
        stats["approved_reports"] = ReportCard.query.filter_by(status="approved").count() if active_term else 0

    # For class teachers — their stream's mark completion
    stream_info = None
    if teacher.stream_id:
        from ..models import Stream, Mark
        stream  = teacher.stream
        students = Student.query.filter_by(stream_id=stream.id, is_active=True).all()

        if active_term and assessments:
            open_assessment = next((a for a in assessments if a.is_open), None)
            if open_assessment:
                marked_students = (
                    db.session.query(Mark.student_id)
                    .filter_by(assessment_id=open_assessment.id)
                    .distinct()
                    .count()
                )
                stream_info = {
                    "stream":           stream,
                    "total_students":   len(students),
                    "marked_students":  marked_students,
                    "open_assessment":  open_assessment,
                }

    return render_template(
        "main/dashboard.html",
        teacher=teacher,
        role=role,
        active_term=active_term,
        assessments=assessments,
        stats=stats,
        stream_info=stream_info,
    )
