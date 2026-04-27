"""
CIS School System - Grading Engine
=====================================
This is the brain of the system.
It handles all grade calculations, ranking, and performance level assignment.
No database calls here — pure logic that can be tested independently.
"""

from .models import (
    GRADING_4BAND,
    GRADING_8BAND,
    get_grade_level,
    get_subjects,
    get_split_subjects,
)


# ---------------------------------------------------------------------------
# SCORE → PERFORMANCE LEVEL
# ---------------------------------------------------------------------------

def assign_performance_level(score, grade_name):
    """
    Given a numeric score and a grade name, return (code, label).

    Junior school (Grade 7–9):
        Score is treated as a percentage (0–100), 8-band scale.

    All other levels:
        Score is raw mark out of 30, 4-band scale.

    Examples:
        assign_performance_level(28, "Grade 3")  → ("EE", "Exceeding Expectation")
        assign_performance_level(87, "Grade 8")  → ("EE2", "Exceeds Expectations")
    """
    if score is None:
        return ("—", "Not Assessed")

    level = get_grade_level(grade_name)
    scale = GRADING_8BAND if level == "junior" else GRADING_4BAND

    for low, high, code, label in scale:
        if low <= round(score) <= high:
            return code, label

    # Fallback — score out of range (shouldn't happen with valid data)
    return ("BE", "Below Expectation")


# ---------------------------------------------------------------------------
# SPLIT SUBJECT COMBINATION
# ---------------------------------------------------------------------------

def combine_split_subject(paper1, paper2, paper1_max, paper2_max):
    """
    Combine two paper scores into a single percentage.

    Example for Junior English (50 + 50):
        combine_split_subject(46, 32, 50, 50) → 78.0

    Example for Upper Primary Kiswahili (40 + 10 insha):
        combine_split_subject(28, 8, 40, 10) → 72.0  (i.e. 36/50 * 100)
    """
    if paper1 is None or paper2 is None:
        return None
    total_max = paper1_max + paper2_max
    if total_max == 0:
        return None
    raw = paper1 + paper2
    return round((raw / total_max) * 100, 2)


# ---------------------------------------------------------------------------
# STUDENT TOTALS & AVERAGES
# ---------------------------------------------------------------------------

def compute_student_summary(student_marks, grade_name):
    """
    Given a list of Mark objects for one student in one assessment,
    return a summary dict with:
        - per-subject breakdown (score, grade_code, grade_label)
        - total score
        - average score
        - performance level
        - subjects_count

    student_marks: list of Mark model instances
    grade_name:    e.g. "Grade 7"
    """
    subjects = get_subjects(grade_name)
    split    = get_split_subjects(grade_name)

    subject_results = {}
    total = 0.0
    count = 0

    for mark in student_marks:
        subj = mark.subject

        # Determine the effective score for this subject
        if subj in split:
            effective = mark.combined_score
        else:
            effective = mark.score

        if effective is None:
            subject_results[subj] = {
                "score": None,
                "paper1": mark.paper1_score,
                "paper2": mark.paper2_score,
                "grade_code": "—",
                "grade_label": "Not Assessed",
            }
            continue

        code, label = assign_performance_level(effective, grade_name)

        subject_results[subj] = {
            "score":       effective,
            "paper1":      mark.paper1_score,
            "paper2":      mark.paper2_score,
            "grade_code":  code,
            "grade_label": label,
        }

        total += effective
        count += 1

    average = round(total / count, 2) if count > 0 else 0.0
    overall_code, overall_label = assign_performance_level(average, grade_name)

    return {
        "subjects":       subject_results,
        "total":          round(total, 2),
        "average":        average,
        "subjects_count": count,
        "overall_code":   overall_code,
        "overall_label":  overall_label,
    }


# ---------------------------------------------------------------------------
# CLASS RANKING
# ---------------------------------------------------------------------------

def rank_students(student_summaries):
    """
    Given a dict of { student_id: summary_dict }, return the same dict
    with a "rank" key added to each summary.

    Ties get the same rank. The next rank after a tie skips numbers
    (standard competition ranking: 1, 2, 2, 4 …).

    student_summaries: { student_id: { "average": float, ... }, ... }
    Returns: same dict with "rank" added
    """
    # Sort by average descending
    sorted_ids = sorted(
        student_summaries.keys(),
        key=lambda sid: student_summaries[sid].get("average", 0),
        reverse=True,
    )

    rank = 1
    for i, sid in enumerate(sorted_ids):
        if i > 0:
            prev_avg = student_summaries[sorted_ids[i - 1]]["average"]
            curr_avg = student_summaries[sid]["average"]
            if curr_avg < prev_avg:
                rank = i + 1  # skip to position (not just +1) to handle ties
        student_summaries[sid]["rank"] = rank

    return student_summaries


# ---------------------------------------------------------------------------
# FULL TERM REPORT COMPUTATION
# ---------------------------------------------------------------------------

def compute_term_report(student, term, db_marks_by_assessment):
    """
    Compute the full-term report for one student covering all 3 assessments.

    Parameters:
        student               : Student model instance
        term                  : Term model instance
        db_marks_by_assessment: dict of { assessment_number: [Mark, ...] }
                                e.g. { 1: [...], 2: [...], 3: [...] }

    Returns a dict:
    {
        "student_name": str,
        "grade":        str,
        "stream":       str,
        "term":         int,
        "year":         int,
        "assessments": {
            1: { subjects: {...}, total, average, overall_code, overall_label },
            2: { ... },
            3: { ... },
        },
        "subjects": [list of subject names for this grade],
        "split_subjects": { subject: (p1_label, p1_max, p2_label, p2_max) },
    }
    """
    grade_name = student.grade.name
    subjects   = get_subjects(grade_name)
    split      = get_split_subjects(grade_name)

    assessments_summary = {}
    for ass_num, marks_list in db_marks_by_assessment.items():
        assessments_summary[ass_num] = compute_student_summary(marks_list, grade_name)

    return {
        "student_name":   student.full_name,
        "admission_no":   student.admission_no,
        "grade":          grade_name,
        "stream":         student.stream.name if student.stream else "",
        "term":           term.term_number,
        "year":           term.academic_year.year,
        "open_date":      term.open_date,
        "close_date":     term.close_date,
        "next_term_date": term.next_term_date,
        "assessments":    assessments_summary,
        "subjects":       subjects,
        "split_subjects": split,
    }


# ---------------------------------------------------------------------------
# GRADE SUMMARY FOR A WHOLE CLASS (used on the admin dashboard)
# ---------------------------------------------------------------------------

def compute_class_summary(grade_name, assessment_marks_dict):
    """
    Compute ranking and averages for an entire class.

    assessment_marks_dict:
        { student_id: [Mark, ...] }  — all marks for one assessment

    Returns:
        { student_id: { "average": float, "rank": int, "subjects": {...} } }
    """
    summaries = {}
    for sid, marks in assessment_marks_dict.items():
        summaries[sid] = compute_student_summary(marks, grade_name)

    return rank_students(summaries)


# ---------------------------------------------------------------------------
# SUBJECT PERFORMANCE ANALYSIS (for admin reports)
# ---------------------------------------------------------------------------

def subject_performance_analysis(grade_name, marks_list):
    """
    Given all marks for a class in one assessment, return per-subject stats:
    mean, highest, lowest, count per performance band.

    Returns:
        { subject: { mean, highest, lowest, bands: { code: count } } }
    """
    from collections import defaultdict

    subject_scores = defaultdict(list)
    for mark in marks_list:
        score = mark.combined_score if mark.combined_score is not None else mark.score
        if score is not None:
            subject_scores[mark.subject].append(score)

    result = {}
    for subj, scores in subject_scores.items():
        bands = {}
        for s in scores:
            code, _ = assign_performance_level(s, grade_name)
            bands[code] = bands.get(code, 0) + 1

        result[subj] = {
            "mean":    round(sum(scores) / len(scores), 2) if scores else 0,
            "highest": max(scores) if scores else 0,
            "lowest":  min(scores) if scores else 0,
            "count":   len(scores),
            "bands":   bands,
        }

    return result
