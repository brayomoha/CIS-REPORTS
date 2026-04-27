"""
CIS School System - Database Models
====================================
Full school hierarchy:
  Reception → PP1 → PP2 → Grade 1–6 → Grade 7–9

Each level group has its own report card format and grading system.
"""

from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

db = SQLAlchemy()


# ===========================================================================
# GRADE LEVEL CONFIG
# ===========================================================================

GRADE_LEVELS = {
    "Reception": "reception",
    "PP1":       "preschool",
    "PP2":       "preschool",
    "Grade 1":   "lower_primary",
    "Grade 2":   "lower_primary",
    "Grade 3":   "lower_primary",
    "Grade 4":   "upper_primary",
    "Grade 5":   "upper_primary",
    "Grade 6":   "upper_primary",
    "Grade 7":   "junior",
    "Grade 8":   "junior",
    "Grade 9":   "junior",
}

# ---------------------------------------------------------------------------
# RECEPTION — skills-based observational report (not marks-based)
# Items rated 1–4: EE=4  ME=3  AE=2  NMP=1 (Needs More Practice)
# RED stream has extra language/writing items (more advanced class)
# ---------------------------------------------------------------------------

RECEPTION_SECTIONS = {
    "Mathematics Activities": [
        "Can rote Count (1 - 10)",
        "Number recognition (1 - 10)",
        "Can recognize all basic shapes",
        "Can recognize colours",
    ],
    "Language Activities": [
        "Express self clearly",
        "Can recite simple poems and sing simple songs",
        "Can recognize all basic pictures",
        "Sound recognition (a - z)",
        "Reading two letter words (card 1)",     # RED stream only
        "Reading Peter and Jane (sight words)",  # RED stream only
    ],
    "Creative Activities": [
        "Can paint",
        "Can colour",
        "Can paste",
        "Music and movement",
    ],
    "Social Skills": [
        "Relate well with peers",
        "Relate well with adults",
        "Is obedient",
    ],
    "Working Skills": [
        "Concentration Skills",
        "Crayon control",
    ],
    "Writing Skills": [
        "Joining dots",                          # RED stream only
    ],
    "Basic Skills": [
        "Feeding",
        "Toileting",
    ],
    "Psychomotor Activities": [
        "Sand play",
        "Running",
        "Jumping",
        "Building blocks",
        "Ballet",
        "Music",
        "Skating",
        "Swimming",
    ],
}

# Items exclusive to RED stream (more advanced)
RECEPTION_RED_ONLY_ITEMS = {
    "Language Activities": [
        "Reading two letter words (card 1)",
        "Reading Peter and Jane (sight words)",
    ],
    "Writing Skills": ["Joining dots"],
}


def get_reception_sections(stream_name="RED"):
    """
    Return section → items dict for the given Reception stream.
    RED stream includes extra language and writing items.
    YELLOW stream uses the simpler item lists.
    """
    is_red = str(stream_name).upper() == "RED"
    result = {}
    for section, items in RECEPTION_SECTIONS.items():
        if not is_red:
            red_only = RECEPTION_RED_ONLY_ITEMS.get(section, [])
            filtered = [i for i in items if i not in red_only]
            if filtered:   # skip Writing Skills entirely for YELLOW (only had 1 item)
                result[section] = filtered
        else:
            result[section] = items
    return result


# ---------------------------------------------------------------------------
# SUBJECTS BY LEVEL GROUP
# ---------------------------------------------------------------------------

SUBJECTS_BY_LEVEL = {
    # Reception sections stored as "subjects" in the marks table
    "reception": list(RECEPTION_SECTIONS.keys()),

    "preschool": [
        "Mathematics Activities",
        "Language Activities",
        "Environmental Activities",
        "Christian Religious Education",
        "Creative Arts",
    ],
    "lower_primary": [
        "Mathematics",
        "English",
        "Kiswahili",
        "Environmental Activities",
        "Christian Religious Education",
        "Creative Arts",
    ],
    "upper_primary": [
        "Mathematics",
        "English",           # split: Paper (40) + Composition (10)
        "Kiswahili",         # split: Paper (40) + Insha (10)
        "Science & Technology",
        "Social Studies",
        "Christian Religious Education",
        "Creative Arts",
        "Agriculture",
    ],
    "junior": [
        "Mathematics",
        "English",           # split: PP1 (50) + PP2 (50)
        "Kiswahili",         # split: PP1 (50) + PP2 (50)
        "Integrated Science",
        "Social Studies",
        "Christian Religious Education",
        "Creative Arts",
        "Agriculture",
        "Pre-Technical Studies",
    ],
}

# ---------------------------------------------------------------------------
# SPLIT SUBJECTS
# Format: { level_group: { subject: (paper1_label, paper1_max, paper2_label, paper2_max) } }
# ---------------------------------------------------------------------------

SPLIT_SUBJECTS = {
    "upper_primary": {
        "English":   ("Paper", 40, "Composition", 10),
        "Kiswahili": ("Paper", 40, "Insha", 10),
    },
    "junior": {
        "English":   ("PP1", 50, "PP2", 50),
        "Kiswahili": ("PP1", 50, "PP2", 50),
    },
}


def get_grade_level(grade_name):
    """Return level group string, e.g. 'Grade 7' → 'junior'"""
    return GRADE_LEVELS.get(grade_name, "lower_primary")


def get_subjects(grade_name):
    """Return subject/section list for a grade"""
    return SUBJECTS_BY_LEVEL.get(get_grade_level(grade_name), [])


def get_split_subjects(grade_name):
    """Return split subject config or empty dict"""
    return SPLIT_SUBJECTS.get(get_grade_level(grade_name), {})


# ===========================================================================
# GRADING SCALES
# ===========================================================================

# Reception: teacher picks level directly (1–4), no numeric score
GRADING_RECEPTION = [
    (4, 4, "EE",  "Exceeding Expectation"),
    (3, 3, "ME",  "Meeting Expectation"),
    (2, 2, "AE",  "Approaching Expectation"),
    (1, 1, "NMP", "Needs More Practice"),
]

# PP1–Grade 6: raw mark out of 30, 4-band
GRADING_4BAND = [
    (26, 30, "EE", "Exceeding Expectation"),
    (18, 25, "ME", "Meeting Expectation"),
    (11, 17, "AE", "Approaching Expectation"),
    (0,  10, "BE", "Below Expectation"),
]

# Grade 7–9: percentage 0–100, 8-band
GRADING_8BAND = [
    (90, 100, "EE1", "Exceeds Expectations"),
    (75,  89, "EE2", "Exceeds Expectations"),
    (58,  74, "ME1", "Meets Expectations"),
    (41,  57, "ME2", "Meets Expectations"),
    (31,  40, "AE1", "Approaches Expectations"),
    (21,  30, "AE2", "Approaches Expectations"),
    (11,  20, "BE1", "Below Expectations"),
    (0,   10, "BE2", "Below Expectations"),
]


def calculate_grade_band(score, grade_name):
    """Return (code, label) for a score at a given grade."""
    if score is None:
        return "—", "Not Assessed"
    level = get_grade_level(grade_name)
    if level == "reception":
        scale = GRADING_RECEPTION
    elif level == "junior":
        scale = GRADING_8BAND
    else:
        scale = GRADING_4BAND
    for low, high, code, label in scale:
        if low <= round(score) <= high:
            return code, label
    return "BE", "Below Expectation"


# ===========================================================================
# DATABASE TABLES
# ===========================================================================

class AcademicYear(db.Model):
    __tablename__ = "academic_years"
    id        = db.Column(db.Integer, primary_key=True)
    year      = db.Column(db.Integer, nullable=False, unique=True)
    is_active = db.Column(db.Boolean, default=True)
    terms     = db.relationship("Term", back_populates="academic_year", cascade="all, delete-orphan")

    def __repr__(self): return f"<AcademicYear {self.year}>"


class Term(db.Model):
    __tablename__ = "terms"
    id               = db.Column(db.Integer, primary_key=True)
    academic_year_id = db.Column(db.Integer, db.ForeignKey("academic_years.id"), nullable=False)
    term_number      = db.Column(db.Integer, nullable=False)
    is_active        = db.Column(db.Boolean, default=False)
    open_date        = db.Column(db.Date, nullable=True)
    close_date       = db.Column(db.Date, nullable=True)
    next_term_date   = db.Column(db.Date, nullable=True)
    academic_year    = db.relationship("AcademicYear", back_populates="terms")
    assessments      = db.relationship("Assessment", back_populates="term", cascade="all, delete-orphan")

    def __repr__(self): return f"<Term {self.term_number} - {self.academic_year.year}>"


class Grade(db.Model):
    """Reception, PP1, PP2, Grade 1 … Grade 9"""
    __tablename__ = "grades"
    id          = db.Column(db.Integer, primary_key=True)
    name        = db.Column(db.String(20), nullable=False, unique=True)
    level_group = db.Column(db.String(20), nullable=False)
    sort_order  = db.Column(db.Integer, nullable=False)
    streams     = db.relationship("Stream", back_populates="grade", cascade="all, delete-orphan")
    students    = db.relationship("Student", back_populates="grade")

    def __repr__(self): return f"<Grade {self.name}>"


class Stream(db.Model):
    """RED or YELLOW stream within a grade"""
    __tablename__ = "streams"
    id       = db.Column(db.Integer, primary_key=True)
    grade_id = db.Column(db.Integer, db.ForeignKey("grades.id"), nullable=False)
    name     = db.Column(db.String(20), nullable=False)
    grade    = db.relationship("Grade", back_populates="streams")
    students = db.relationship("Student", back_populates="stream")
    teachers = db.relationship("Teacher", back_populates="stream")

    def __repr__(self): return f"<Stream {self.grade.name} {self.name}>"


class Student(db.Model):
    __tablename__ = "students"
    id           = db.Column(db.Integer, primary_key=True)
    admission_no = db.Column(db.String(20), unique=True, nullable=False)
    full_name    = db.Column(db.String(100), nullable=False)
    grade_id     = db.Column(db.Integer, db.ForeignKey("grades.id"), nullable=False)
    stream_id    = db.Column(db.Integer, db.ForeignKey("streams.id"), nullable=True)
    gender       = db.Column(db.String(10), nullable=True)
    parent_email = db.Column(db.String(120), nullable=True)
    is_active    = db.Column(db.Boolean, default=True)
    enrolled_on  = db.Column(db.Date, default=datetime.utcnow)
    grade        = db.relationship("Grade", back_populates="students")
    stream       = db.relationship("Stream", back_populates="students")
    marks        = db.relationship("Mark", back_populates="student", cascade="all, delete-orphan")
    reports      = db.relationship("ReportCard", back_populates="student", cascade="all, delete-orphan")
    skill_ratings = db.relationship("SkillRating", back_populates="student", cascade="all, delete-orphan")

    def __repr__(self): return f"<Student {self.full_name}>"


class Teacher(db.Model):
    __tablename__ = "teachers"
    id            = db.Column(db.Integer, primary_key=True)
    full_name     = db.Column(db.String(100), nullable=False)
    email         = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    role          = db.Column(db.String(20), default="teacher")
    is_active     = db.Column(db.Boolean, default=True)
    stream_id     = db.Column(db.Integer, db.ForeignKey("streams.id"), nullable=True)
    stream        = db.relationship("Stream", back_populates="teachers")
    # Subject restriction: blank = all subjects, comma-separated = specific subjects
    # e.g. "Mathematics" or "Mathematics,Science & Technology"
    subject       = db.Column(db.String(200), nullable=True)
    # Extra stream IDs stored as comma-separated string e.g. "3,7,12"
    extra_streams = db.Column(db.String(200), nullable=True)

    def get_all_stream_ids(self):
        """All stream IDs this teacher can access."""
        ids = set()
        if self.stream_id:
            ids.add(self.stream_id)
        if self.extra_streams:
            for s in self.extra_streams.split(","):
                s = s.strip()
                if s.isdigit():
                    ids.add(int(s))
        return ids

    def can_access_stream(self, stream_id):
        if self.role in ("admin", "principal"):
            return True
        return stream_id in self.get_all_stream_ids()

    def can_enter_subject(self, subject_name):
        if self.role in ("admin", "principal"):
            return True
        if not self.subject:
            return True  # no restriction — all subjects
        # Support multiple subjects separated by comma
        allowed = [s.strip().lower() for s in self.subject.split(",")]
        return subject_name.lower() in allowed

    def __repr__(self): return f"<Teacher {self.full_name} [{self.role}]>"


class Assessment(db.Model):
    """Entry Assessment (1), Mid Term (2), End Term (3). Reception skips Entry."""
    __tablename__ = "assessments"
    id      = db.Column(db.Integer, primary_key=True)
    term_id = db.Column(db.Integer, db.ForeignKey("terms.id"), nullable=False)
    name    = db.Column(db.String(30), nullable=False)
    number  = db.Column(db.Integer, nullable=False)
    is_open = db.Column(db.Boolean, default=False)
    term    = db.relationship("Term", back_populates="assessments")
    marks   = db.relationship("Mark", back_populates="assessment", cascade="all, delete-orphan")

    def __repr__(self): return f"<Assessment {self.name}>"


class Mark(db.Model):
    """
    One subject score for one student in one assessment.

    Reception:      subject = section name, score = None (use SkillRating table instead)
    Split subjects: paper1_score + paper2_score → combined_score (%)
    Others:         score = raw mark
    """
    __tablename__ = "marks"
    id            = db.Column(db.Integer, primary_key=True)
    student_id    = db.Column(db.Integer, db.ForeignKey("students.id"), nullable=False)
    assessment_id = db.Column(db.Integer, db.ForeignKey("assessments.id"), nullable=False)
    subject       = db.Column(db.String(80), nullable=False)
    score          = db.Column(db.Float, nullable=True)
    paper1_score   = db.Column(db.Float, nullable=True)
    paper2_score   = db.Column(db.Float, nullable=True)
    combined_score = db.Column(db.Float, nullable=True)
    grade_code     = db.Column(db.String(5), nullable=True)
    grade_label    = db.Column(db.String(40), nullable=True)
    entered_by     = db.Column(db.Integer, db.ForeignKey("teachers.id"), nullable=True)
    entered_at     = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at     = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    __table_args__ = (
        db.UniqueConstraint("student_id", "assessment_id", "subject",
                            name="uq_student_assessment_subject"),
    )
    student    = db.relationship("Student", back_populates="marks")
    assessment = db.relationship("Assessment", back_populates="marks")

    def effective_score(self):
        return self.combined_score if self.combined_score is not None else self.score

    def __repr__(self): return f"<Mark {self.subject} | {self.effective_score()}>"


class SkillRating(db.Model):
    """
    Individual skill item rating for Reception learners.
    rating: 1=NMP  2=AE  3=ME  4=EE
    """
    __tablename__ = "skill_ratings"
    id            = db.Column(db.Integer, primary_key=True)
    student_id    = db.Column(db.Integer, db.ForeignKey("students.id"), nullable=False)
    assessment_id = db.Column(db.Integer, db.ForeignKey("assessments.id"), nullable=False)
    section       = db.Column(db.String(60), nullable=False)    # e.g. "Mathematics Activities"
    skill_item    = db.Column(db.String(120), nullable=False)   # e.g. "Can rote Count (1 - 10)"
    rating        = db.Column(db.Integer, nullable=True)         # 1, 2, 3, or 4
    __table_args__ = (
        db.UniqueConstraint("student_id", "assessment_id", "skill_item",
                            name="uq_student_assessment_skill"),
    )
    student    = db.relationship("Student", back_populates="skill_ratings")
    assessment = db.relationship("Assessment")

    @property
    def code(self):
        return {4: "EE", 3: "ME", 2: "AE", 1: "NMP"}.get(self.rating, "—")

    @property
    def label(self):
        return {
            4: "Exceeding Expectation",
            3: "Meeting Expectation",
            2: "Approaching Expectation",
            1: "Needs More Practice",
        }.get(self.rating, "Not Assessed")

    def __repr__(self): return f"<SkillRating {self.skill_item} | {self.code}>"


class ReportCard(db.Model):
    """
    Completed report card for one student for a full term.
    Reception uses general_comment (one block).
    PP1 and above use three CBC comment sections.
    """
    __tablename__ = "report_cards"
    id         = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.Integer, db.ForeignKey("students.id"), nullable=False)
    term_id    = db.Column(db.Integer, db.ForeignKey("terms.id"), nullable=False)

    # CBC three-part comments (PP1 and above)
    comment_performance  = db.Column(db.Text, nullable=True)
    comment_competencies = db.Column(db.Text, nullable=True)
    comment_values       = db.Column(db.Text, nullable=True)

    # Reception single comment
    general_comment = db.Column(db.Text, nullable=True)

    # Workflow: draft → comments_pending → pending_approval → approved → sent
    status      = db.Column(db.String(20), default="draft")
    class_rank  = db.Column(db.Integer, nullable=True)
    overall_avg = db.Column(db.Float, nullable=True)

    pdf_path     = db.Column(db.String(256), nullable=True)
    generated_at = db.Column(db.DateTime, nullable=True)
    approved_by  = db.Column(db.Integer, db.ForeignKey("teachers.id"), nullable=True)
    approved_at  = db.Column(db.DateTime, nullable=True)
    emailed_at   = db.Column(db.DateTime, nullable=True)

    __table_args__ = (
        db.UniqueConstraint("student_id", "term_id", name="uq_student_term_report"),
    )
    student = db.relationship("Student", back_populates="reports")
    term    = db.relationship("Term")

    def __repr__(self): return f"<ReportCard {self.student.full_name} | Term {self.term.term_number}>"
