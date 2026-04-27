"""
CIS School System - Marks Entry Routes
========================================
Teachers use these pages to enter subject marks for their class.
Handles both single-paper and split-paper (English/Kiswahili) subjects.
"""

from flask import Blueprint, render_template, request, redirect, url_for, flash, session, jsonify
from ..models import db, Teacher, Student, Grade, Stream, Assessment, Mark, Term
from ..models import get_subjects, get_split_subjects, calculate_grade_band
from ..grading import combine_split_subject, assign_performance_level
from .auth import login_required

marks_bp = Blueprint("marks", __name__)


@marks_bp.route("/")
@login_required
def index():
    """List all grade/stream/assessment combinations the teacher can enter marks for"""
    teacher = Teacher.query.get(session["teacher_id"])
    active_term = Term.query.filter_by(is_active=True).first()

    if not active_term:
        flash("No active term found. Contact admin.", "warning")
        return redirect(url_for("main.dashboard"))

    open_assessments = Assessment.query.filter_by(term_id=active_term.id, is_open=True).all()

    if teacher.role in ("admin", "principal"):
        grades  = Grade.query.order_by(Grade.sort_order).all()
        streams = Stream.query.join(Grade).order_by(Grade.sort_order, Stream.name).all()
    else:
        all_ids = teacher.get_all_stream_ids()
        if not all_ids:
            flash("You are not assigned to any class. Contact admin.", "warning")
            return redirect(url_for("main.dashboard"))
        streams = Stream.query.filter(Stream.id.in_(all_ids))                      .join(Grade).order_by(Grade.sort_order, Stream.name).all()
        grades  = list({s.grade for s in streams})
        grades.sort(key=lambda g: g.sort_order)

    return render_template(
        "marks/index.html",
        teacher=teacher,
        grades=grades,
        streams=streams,
        open_assessments=open_assessments,
        active_term=active_term,
    )


@marks_bp.route("/enter/<int:assessment_id>/<int:stream_id>")
@login_required
def enter_marks(assessment_id, stream_id):
    """Show the mark entry sheet for a specific assessment and stream"""
    teacher    = Teacher.query.get(session["teacher_id"])
    assessment = Assessment.query.get_or_404(assessment_id)
    stream     = Stream.query.get_or_404(stream_id)
    grade      = stream.grade

    # Security: non-admin teachers can only enter marks for their own stream
    if teacher.role not in ("admin", "principal") and not teacher.can_access_stream(stream_id):
        flash("You can only enter marks for your assigned class.", "danger")
        return redirect(url_for("marks.index"))

    if not assessment.is_open:
        flash("This assessment is not open for mark entry.", "warning")
        return redirect(url_for("marks.index"))

    students = (
        Student.query
        .filter_by(stream_id=stream_id, is_active=True)
        .order_by(Student.full_name)
        .all()
    )

    subjects      = get_subjects(grade.name)
    split_subjects = get_split_subjects(grade.name)

    # Load existing marks for this assessment + stream
    student_ids = [s.id for s in students]
    existing_marks = (
        Mark.query
        .filter(
            Mark.assessment_id == assessment_id,
            Mark.student_id.in_(student_ids)
        )
        .all()
    )

    # Build a lookup: { student_id: { subject: Mark } }
    marks_lookup = {}
    for mark in existing_marks:
        marks_lookup.setdefault(mark.student_id, {})[mark.subject] = mark

    from ..models import get_grade_level
    level     = get_grade_level(grade.name)
    max_score = 100 if level == "junior" else 30

    # Filter subjects based on teacher assignment
    # Subject teachers only see their subject; class teachers see all
    if teacher.role in ("admin", "principal") or not teacher.subject:
        allowed_subjects = subjects
        teacher_subject  = None
    else:
        allowed_subjects = [s for s in subjects if teacher.can_enter_subject(s)]
        teacher_subject  = teacher.subject

    return render_template(
        "marks/enter.html",
        teacher=teacher,
        assessment=assessment,
        stream=stream,
        grade=grade,
        students=students,
        subjects=subjects,
        allowed_subjects=allowed_subjects,
        teacher_subject=teacher_subject,
        split_subjects=split_subjects,
        marks_lookup=marks_lookup,
        max_score=max_score,
    )


@marks_bp.route("/save", methods=["POST"])
@login_required
def save_marks():
    """
    Save submitted marks to the database.
    Called when teacher submits the mark entry form.
    Handles both single-paper and split-paper subjects.
    """
    teacher       = Teacher.query.get(session["teacher_id"])
    assessment_id = int(request.form.get("assessment_id"))
    stream_id     = int(request.form.get("stream_id"))
    assessment    = Assessment.query.get_or_404(assessment_id)
    stream        = Stream.query.get_or_404(stream_id)
    grade         = stream.grade

    if teacher.role not in ("admin", "principal") and not teacher.can_access_stream(stream_id):
        flash("Permission denied.", "danger")
        return redirect(url_for("marks.index"))

    if not assessment.is_open:
        flash("This assessment is closed.", "warning")
        return redirect(url_for("marks.index"))

    subjects       = get_subjects(grade.name)
    split_subjects = get_split_subjects(grade.name)

    students = Student.query.filter_by(stream_id=stream_id, is_active=True).all()
    saved    = 0

    for student in students:
        for subject in subjects:
            # Build form field key (spaces replaced with underscores for HTML)
            field_key = f"{student.id}_{subject.replace(' ', '_').replace('&', 'and')}"

            # All subjects now use a single score input (no split papers)
            raw = request.form.get(field_key, "").strip()
            # Store as whole integer — no decimals allowed
            score     = int(round(float(raw))) if raw else None
            paper1    = None
            paper2    = None
            combined  = None
            effective = score

            # Grade band
            if effective is not None:
                code, label = assign_performance_level(effective, grade.name)
            else:
                code, label = None, None

            # Skip empty inputs entirely — don't save None marks
            if effective is None:
                continue

            # Upsert — update if exists, insert if not
            mark = Mark.query.filter_by(
                student_id=student.id,
                assessment_id=assessment_id,
                subject=subject,
            ).first()

            if mark is None:
                mark = Mark(
                    student_id=student.id,
                    assessment_id=assessment_id,
                    subject=subject,
                    entered_by=teacher.id,
                )
                db.session.add(mark)

            # Save as single score for all subjects
            mark.score          = effective
            mark.paper1_score   = None
            mark.paper2_score   = None
            mark.combined_score = None

            mark.grade_code  = code
            mark.grade_label = label
            saved += 1

    db.session.commit()
    flash(f"✅ Marks saved successfully for {stream.grade.name} {stream.name} — {assessment.name}.", "success")
    return redirect(url_for("marks.enter_marks", assessment_id=assessment_id, stream_id=stream_id))


@marks_bp.route("/api/student/<int:student_id>/marks/<int:assessment_id>")
@login_required
def get_student_marks(student_id, assessment_id):
    """API endpoint — returns marks for one student as JSON (used by dashboard widgets)"""
    marks = Mark.query.filter_by(student_id=student_id, assessment_id=assessment_id).all()
    return jsonify([
        {
            "subject":         m.subject,
            "score":           m.score,
            "paper1":          m.paper1_score,
            "paper2":          m.paper2_score,
            "combined":        m.combined_score,
            "grade_code":      m.grade_code,
            "grade_label":     m.grade_label,
        }
        for m in marks
    ])
