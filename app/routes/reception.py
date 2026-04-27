"""
CIS School System - Reception Marks Entry Routes
=================================================
Reception uses a skills-based rating system (1–4) instead of numeric marks.
Each skill item is rated individually per student.
RED and YELLOW streams have different item sets.
"""

from flask import Blueprint, render_template, request, redirect, url_for, flash, session
from ..models import (
    db, Teacher, Student, Stream, Assessment, Term, SkillRating, ReportCard,
    get_reception_sections
)
from .auth import login_required

reception_bp = Blueprint("reception", __name__)

RATING_LABELS = {
    4: ("EE",  "Exceeding Expectation"),
    3: ("ME",  "Meeting Expectation"),
    2: ("AE",  "Approaching Expectation"),
    1: ("NMP", "Needs More Practice"),
}


@reception_bp.route("/enter/<int:assessment_id>/<int:stream_id>")
@login_required
def enter_skills(assessment_id, stream_id):
    """Show the Reception skill rating sheet for a stream"""
    teacher    = Teacher.query.get(session["teacher_id"])
    assessment = Assessment.query.get_or_404(assessment_id)
    stream     = Stream.query.get_or_404(stream_id)

    # Security
    if teacher.role not in ("admin", "principal") and teacher.stream_id != stream_id:
        flash("You can only enter ratings for your assigned class.", "danger")
        return redirect(url_for("marks.index"))

    if not assessment.is_open:
        flash("This assessment is not open for entry.", "warning")
        return redirect(url_for("marks.index"))

    students = (
        Student.query
        .filter_by(stream_id=stream_id, is_active=True)
        .order_by(Student.full_name)
        .all()
    )

    # Get sections appropriate for this stream (RED vs YELLOW)
    sections = get_reception_sections(stream.name)

    # Load existing ratings
    student_ids = [s.id for s in students]
    existing = SkillRating.query.filter(
        SkillRating.assessment_id == assessment_id,
        SkillRating.student_id.in_(student_ids)
    ).all()

    # Build lookup: { student_id: { skill_item: rating } }
    ratings_lookup = {}
    for r in existing:
        ratings_lookup.setdefault(r.student_id, {})[r.skill_item] = r.rating

    return render_template(
        "reception/enter_skills.html",
        teacher=teacher,
        assessment=assessment,
        stream=stream,
        students=students,
        sections=sections,
        ratings_lookup=ratings_lookup,
        rating_labels=RATING_LABELS,
    )


@reception_bp.route("/save", methods=["POST"])
@login_required
def save_skills():
    """Save submitted skill ratings to the database"""
    teacher       = Teacher.query.get(session["teacher_id"])
    assessment_id = int(request.form.get("assessment_id"))
    stream_id     = int(request.form.get("stream_id"))
    assessment    = Assessment.query.get_or_404(assessment_id)
    stream        = Stream.query.get_or_404(stream_id)

    if teacher.role not in ("admin", "principal") and teacher.stream_id != stream_id:
        flash("Permission denied.", "danger")
        return redirect(url_for("marks.index"))

    if not assessment.is_open:
        flash("This assessment is closed.", "warning")
        return redirect(url_for("marks.index"))

    sections = get_reception_sections(stream.name)
    students = Student.query.filter_by(stream_id=stream_id, is_active=True).all()

    for student in students:
        for section, items in sections.items():
            for item in items:
                # Form field: student_id__item_name (spaces → underscores)
                field_key = f"{student.id}__{item.replace(' ', '_').replace('(', '').replace(')', '').replace('-', '').replace('/', '')}"
                raw = request.form.get(field_key, "").strip()
                rating = int(raw) if raw and raw.isdigit() else None

                # Upsert
                sr = SkillRating.query.filter_by(
                    student_id=student.id,
                    assessment_id=assessment_id,
                    skill_item=item,
                ).first()

                if sr is None:
                    sr = SkillRating(
                        student_id=student.id,
                        assessment_id=assessment_id,
                        section=section,
                        skill_item=item,
                    )
                    db.session.add(sr)

                sr.rating = rating

    db.session.commit()
    flash(f"✅ Skill ratings saved for Reception {stream.name} — {assessment.name}.", "success")
    return redirect(url_for(
        "reception.enter_skills",
        assessment_id=assessment_id,
        stream_id=stream_id
    ))


@reception_bp.route("/comments/<int:student_id>/edit", methods=["GET", "POST"])
@login_required
def edit_comment(student_id):
    """Edit the single general comment for a Reception learner"""
    teacher     = Teacher.query.get(session["teacher_id"])
    student     = Student.query.get_or_404(student_id)
    active_term = Term.query.filter_by(is_active=True).first()

    if teacher.role not in ("admin", "principal") and teacher.stream_id != student.stream_id:
        flash("Permission denied.", "danger")
        return redirect(url_for("main.dashboard"))

    rc = ReportCard.query.filter_by(student_id=student_id, term_id=active_term.id).first()
    if not rc:
        rc = ReportCard(student_id=student_id, term_id=active_term.id, status="draft")
        db.session.add(rc)
        db.session.commit()

    if request.method == "POST":
        rc.general_comment = request.form.get("general_comment", "").strip()
        rc.status = "pending_approval" if rc.general_comment else "comments_pending"
        db.session.commit()
        flash(f"Comment saved for {student.full_name}.", "success")

        if request.form.get("action") == "save_next":
            students = (
                Student.query
                .filter_by(stream_id=student.stream_id, is_active=True)
                .order_by(Student.full_name)
                .all()
            )
            ids = [s.id for s in students]
            idx = ids.index(student_id)
            if idx + 1 < len(ids):
                return redirect(url_for("reception.edit_comment", student_id=ids[idx + 1]))

        return redirect(url_for("reports.comments_list", stream_id=student.stream_id))

    # Load skill ratings for the preview
    assessments = Assessment.query.filter_by(term_id=active_term.id).order_by(Assessment.number).all()
    skills_by_assessment = {}
    sections = get_reception_sections(student.stream.name if student.stream else "RED")

    for ass in assessments:
        if ass.number == 1:
            continue  # Reception skips Entry Assessment
        ratings = SkillRating.query.filter_by(
            student_id=student_id,
            assessment_id=ass.id
        ).all()
        skills_by_assessment[ass.number] = {r.skill_item: r for r in ratings}

    return render_template(
        "reception/edit_comment.html",
        student=student,
        rc=rc,
        sections=sections,
        skills_by_assessment=skills_by_assessment,
        rating_labels=RATING_LABELS,
        teacher=teacher,
        active_term=active_term,
    )
