"""
CIS School System - Comments Bank
===================================
A built-in library of pre-written CBC comments organised by:
  - Level group (preschool, lower_primary, upper_primary, junior, reception)
  - Comment type (performance, competencies, values, general)
  - Performance band (EE, ME, AE, BE / NMP)

This replicates the Google Sheet comments file inside the system so teachers
can either pick a pre-written comment or type their own.

The database table (CommentTemplate) stores all comments and allows admin
to add, edit, or import more from Excel/Google Sheets at any time.
"""

from .models import db


class CommentTemplate(db.Model):
    """
    A pre-written comment template teachers can select from.
    Stored in the database so admin can manage them via the UI.
    """
    __tablename__ = "comment_templates"

    id           = db.Column(db.Integer, primary_key=True)
    level_group  = db.Column(db.String(20), nullable=False)   # reception / preschool / lower_primary / upper_primary / junior
    comment_type = db.Column(db.String(20), nullable=False)   # performance / competencies / values / general
    band         = db.Column(db.String(5),  nullable=True)    # EE / ME / AE / BE / NMP / None (applies to all)
    gender       = db.Column(db.String(10), nullable=True)    # "Male" / "Female" / None (gender-neutral)
    text         = db.Column(db.Text, nullable=False)
    is_active    = db.Column(db.Boolean, default=True)

    def __repr__(self):
        return f"<CommentTemplate [{self.level_group}|{self.comment_type}|{self.band}]>"


# ---------------------------------------------------------------------------
# SEED DATA — starter comments for each level
# These mirror the structure of your Google Sheet.
# Admin can add/edit more via the UI, or import from Excel.
# ---------------------------------------------------------------------------

STARTER_COMMENTS = [

    # -----------------------------------------------------------------------
    # RECEPTION — General Comments
    # -----------------------------------------------------------------------
    ("reception", "general", "EE", None,
     "{name} is a bright and enthusiastic learner who has shown remarkable progress this term. "
     "They participate actively in all activities and demonstrate excellent understanding of concepts taught."),
    ("reception", "general", "EE", None,
     "{name} has had an outstanding term. They engage confidently with all activities, show strong "
     "social skills, and consistently demonstrate a love for learning."),
    ("reception", "general", "ME", None,
     "{name} has had a good term and is making steady progress. They participate well in class "
     "activities and are developing their skills in all areas. With continued practice, {pronoun} will "
     "continue to grow."),
    ("reception", "general", "ME", None,
     "{name} is a pleasant learner who follows instructions well. {pronoun_cap} is making satisfactory "
     "progress and shows a positive attitude towards school activities."),
    ("reception", "general", "AE", None,
     "{name} is making gradual progress this term. {pronoun_cap} needs more practice in some areas "
     "and benefits from additional encouragement and support. We look forward to seeing continued "
     "improvement next term."),
    ("reception", "general", "NMP", None,
     "{name} requires more support and practice to develop foundational skills. We encourage regular "
     "practice at home and look forward to working with {pronoun} to help {pronoun} reach {pronoun_pos} "
     "full potential."),

    # -----------------------------------------------------------------------
    # PRESCHOOL (PP1 & PP2) — Performance
    # -----------------------------------------------------------------------
    ("preschool", "performance", "EE", None,
     "{name} has had an excellent term. {pronoun_cap} demonstrates a strong grasp of all learning "
     "areas and consistently produces outstanding work. {pronoun_cap} is a joy to teach."),
    ("preschool", "performance", "EE", None,
     "{name} is a highly motivated learner who shows exceptional understanding across all activities. "
     "{pronoun_cap} completes tasks with accuracy and enthusiasm."),
    ("preschool", "performance", "ME", None,
     "{name} has performed well this term and is meeting expectations in all learning areas. "
     "{pronoun_cap} shows good understanding and participates positively in class activities."),
    ("preschool", "performance", "ME", None,
     "{name} is making good progress and demonstrates satisfactory understanding of the concepts "
     "taught this term. With continued effort, {pronoun} will continue to improve."),
    ("preschool", "performance", "AE", None,
     "{name} is working towards meeting expectations. {pronoun_cap} requires additional support in "
     "some learning areas and will benefit from regular practice and revision at home."),
    ("preschool", "performance", "BE", None,
     "{name} needs significant support to grasp the concepts taught this term. We encourage more "
     "practice at home and ask that parents/guardians work closely with the teacher to support "
     "{pronoun_pos} learning."),

    # PRESCHOOL — Competencies
    ("preschool", "competencies", None, None,
     "{pronoun_cap} participates actively during learning activities, which promotes communication and collaboration skills."),
    ("preschool", "competencies", None, None,
     "{pronoun_cap} demonstrates creativity during art and play activities, showing imagination and self-expression."),
    ("preschool", "competencies", None, None,
     "{pronoun_cap} is developing good self-management skills by following routines and completing tasks independently."),
    ("preschool", "competencies", None, None,
     "{pronoun_cap} engages well with peers during group activities, demonstrating early collaboration and social skills."),

    # PRESCHOOL — Values
    ("preschool", "values", None, None,
     "{pronoun_cap} is a respectful and well-behaved learner who follows instructions appropriately."),
    ("preschool", "values", None, None,
     "{pronoun_cap} demonstrates honesty and kindness in {pronoun_pos} interactions with peers and adults."),
    ("preschool", "values", None, None,
     "{pronoun_cap} is a disciplined learner who treats others with respect and follows school rules consistently."),

    # -----------------------------------------------------------------------
    # LOWER PRIMARY (Grade 1–3) — Performance
    # -----------------------------------------------------------------------
    ("lower_primary", "performance", "EE", None,
     "{name} has had an exceptional term, consistently performing at the highest level across all "
     "learning areas. {pronoun_cap} demonstrates excellent understanding, completes work accurately, "
     "and shows a genuine love for learning."),
    ("lower_primary", "performance", "EE", None,
     "{name} is a remarkable learner who has impressed this term with {pronoun_pos} dedication and "
     "high performance. {pronoun_cap} grasps concepts quickly and applies them effectively."),
    ("lower_primary", "performance", "ME", None,
     "{name} has had a good term and is meeting expectations across all learning areas. "
     "{pronoun_cap} works diligently and shows a positive attitude towards school work."),
    ("lower_primary", "performance", "ME", None,
     "{name} demonstrates satisfactory understanding of the concepts taught this term. "
     "{pronoun_cap} participates well and is making steady progress. Continued effort will help "
     "{pronoun} achieve even better results."),
    ("lower_primary", "performance", "AE", None,
     "{name} is working towards meeting expectations and requires additional practice in some areas. "
     "{pronoun_cap} would benefit from regular reading and revision at home."),
    ("lower_primary", "performance", "BE", None,
     "{name} is finding some concepts challenging this term and requires targeted support. "
     "We encourage parents to work with {pronoun} at home and liaise closely with the class teacher."),

    # LOWER PRIMARY — Competencies
    ("lower_primary", "competencies", None, None,
     "{pronoun_cap} communicates clearly and confidently during class discussions, demonstrating strong communication skills."),
    ("lower_primary", "competencies", None, None,
     "{pronoun_cap} shows creativity in written and art work, expressing {pronoun_pos} ideas imaginatively."),
    ("lower_primary", "competencies", None, None,
     "{pronoun_cap} works well in group activities, showing the ability to collaborate and share ideas with peers."),
    ("lower_primary", "competencies", None, None,
     "{pronoun_cap} is developing good critical thinking skills by asking thoughtful questions and engaging with learning tasks."),
    ("lower_primary", "competencies", None, None,
     "{pronoun_cap} demonstrates self-management by completing assignments on time and maintaining an organised workbook."),

    # LOWER PRIMARY — Values
    ("lower_primary", "values", None, None,
     "{pronoun_cap} is a respectful and responsible learner who follows school rules consistently."),
    ("lower_primary", "values", None, None,
     "{pronoun_cap} demonstrates integrity and honesty in {pronoun_pos} daily interactions with peers and teachers."),
    ("lower_primary", "values", None, None,
     "{pronoun_cap} is kind and considerate to classmates, contributing to a positive classroom environment."),

    # -----------------------------------------------------------------------
    # UPPER PRIMARY (Grade 4–6) — Performance
    # -----------------------------------------------------------------------
    ("upper_primary", "performance", "EE", None,
     "{name} has had an outstanding term, achieving at the highest level in all subjects. "
     "{pronoun_cap} demonstrates excellent analytical skills, produces high quality work, and is "
     "a consistent top performer in the class."),
    ("upper_primary", "performance", "EE", None,
     "{name} continues to excel academically. {pronoun_pos} depth of understanding, attention to "
     "detail, and commitment to excellence make {pronoun} a model learner."),
    ("upper_primary", "performance", "ME", None,
     "{name} has performed well this term and is meeting expectations across all subjects. "
     "{pronoun_cap} demonstrates good understanding and submits work consistently. "
     "With continued focus, {pronoun} can push for even higher results."),
    ("upper_primary", "performance", "ME", None,
     "{name} shows satisfactory performance this term. {pronoun_cap} participates actively in "
     "lessons and demonstrates a willingness to learn. Consistent revision and practice will "
     "strengthen {pronoun_pos} results."),
    ("upper_primary", "performance", "AE", None,
     "{name} is approaching expectations in several subjects and will benefit from more structured "
     "revision and exam practice. We encourage {pronoun} to seek help when concepts are unclear."),
    ("upper_primary", "performance", "BE", None,
     "{name} requires significant support this term. We recommend additional tuition and consistent "
     "study at home. The school will continue to provide support, and we ask parents to partner "
     "with us in this process."),

    # UPPER PRIMARY — Competencies
    ("upper_primary", "competencies", None, None,
     "{pronoun_cap} demonstrates strong communication skills through well-structured written work and confident classroom participation."),
    ("upper_primary", "competencies", None, None,
     "{pronoun_cap} shows excellent critical thinking by analysing problems carefully and arriving at logical conclusions."),
    ("upper_primary", "competencies", None, None,
     "{pronoun_cap} collaborates effectively with peers during group work, contributing meaningfully and respecting different perspectives."),
    ("upper_primary", "competencies", None, None,
     "{pronoun_cap} is developing strong digital literacy skills and applies technology appropriately in learning tasks."),
    ("upper_primary", "competencies", None, None,
     "{pronoun_cap} demonstrates creativity and imagination in project work, often going beyond the required standard."),

    # UPPER PRIMARY — Values
    ("upper_primary", "values", None, None,
     "{pronoun_cap} is a responsible and respectful learner who upholds the school's values consistently."),
    ("upper_primary", "values", None, None,
     "{pronoun_cap} demonstrates patriotism and social responsibility by participating in school and community activities."),
    ("upper_primary", "values", None, None,
     "{pronoun_cap} shows integrity in {pronoun_pos} work and dealings with others, consistently choosing to do the right thing."),

    # -----------------------------------------------------------------------
    # JUNIOR SCHOOL (Grade 7–9) — Performance
    # -----------------------------------------------------------------------
    ("junior", "performance", "EE", None,
     "{name} has had an exceptional term, excelling across all subjects and demonstrating a high "
     "level of academic maturity. {pronoun_cap} produces well-researched, insightful work and is "
     "a role model for {pronoun_pos} peers."),
    ("junior", "performance", "EE", None,
     "{name} continues to demonstrate outstanding academic ability. {pronoun_pos} work is "
     "consistently thorough and of excellent quality. We commend {pronoun} for {pronoun_pos} "
     "commitment and hard work."),
    ("junior", "performance", "ME", None,
     "{name} has performed commendably this term, meeting expectations across all subjects. "
     "{pronoun_cap} participates actively in discussions and demonstrates a solid understanding "
     "of the curriculum. With greater focus on weaker areas, {pronoun} can achieve even more."),
    ("junior", "performance", "ME", None,
     "{name} shows satisfactory performance this term. {pronoun_cap} is a dedicated learner who "
     "submits work on time and engages positively in class. Consistent revision will help "
     "{pronoun} improve further."),
    ("junior", "performance", "AE", None,
     "{name} is approaching expectations and has room for significant improvement. "
     "{pronoun_cap} should prioritise regular study, past paper practice, and seek clarification "
     "on topics that are unclear. We believe {pronoun} can do better."),
    ("junior", "performance", "BE", None,
     "{name} has found this term challenging and requires intensive support. We strongly encourage "
     "additional tutoring, disciplined study habits, and close collaboration between home and school "
     "to help {pronoun} meet expectations."),

    # JUNIOR — Competencies
    ("junior", "competencies", None, None,
     "{pronoun_cap} demonstrates excellent communication skills through articulate written and verbal contributions in class."),
    ("junior", "competencies", None, None,
     "{pronoun_cap} applies critical thinking and problem-solving skills effectively, particularly in Mathematics and Sciences."),
    ("junior", "competencies", None, None,
     "{pronoun_cap} collaborates well in group settings, demonstrating leadership and the ability to work towards shared goals."),
    ("junior", "competencies", None, None,
     "{pronoun_cap} shows strong self-management skills — arriving prepared for lessons, meeting deadlines, and maintaining focus."),
    ("junior", "competencies", None, None,
     "{pronoun_cap} demonstrates digital literacy and uses technology responsibly to enhance learning."),
    ("junior", "competencies", None, None,
     "{pronoun_cap} shows creativity and innovation in project work, consistently bringing fresh perspectives to assignments."),

    # JUNIOR — Values
    ("junior", "values", None, None,
     "{pronoun_cap} consistently upholds the school's values of respect, responsibility, and integrity."),
    ("junior", "values", None, None,
     "{pronoun_cap} demonstrates a strong sense of social responsibility and contributes positively to the school community."),
    ("junior", "values", None, None,
     "{pronoun_cap} shows patriotism and a commitment to excellence that reflects well on the school and family."),
    ("junior", "values", None, None,
     "{pronoun_cap} treats peers and teachers with respect and demonstrates empathy in {pronoun_pos} daily interactions."),
]


def fill_comment(template_text, student_name, gender="unknown"):
    """
    Replace placeholders in a comment template with student-specific values.

    Placeholders:
        {name}         → student's full name
        {pronoun}      → he / she / they
        {pronoun_cap}  → He / She / They
        {pronoun_pos}  → his / her / their
    """
    gender = (gender or "unknown").lower()
    if gender == "male":
        pronoun, pronoun_cap, pronoun_pos = "he", "He", "his"
    elif gender == "female":
        pronoun, pronoun_cap, pronoun_pos = "she", "She", "her"
    else:
        pronoun, pronoun_cap, pronoun_pos = "they", "They", "their"

    return (
        template_text
        .replace("{name}", student_name)
        .replace("{pronoun_cap}", pronoun_cap)
        .replace("{pronoun_pos}", pronoun_pos)
        .replace("{pronoun}", pronoun)
    )


def seed_comment_templates(app):
    """
    Populate the comment_templates table with starter comments.
    Safe to call multiple times — skips if already populated.
    """
    with app.app_context():
        if CommentTemplate.query.count() > 0:
            return  # already seeded

        for level_group, comment_type, band, gender, text in STARTER_COMMENTS:
            db.session.add(CommentTemplate(
                level_group=level_group,
                comment_type=comment_type,
                band=band,
                gender=gender,
                text=text,
                is_active=True,
            ))
        db.session.commit()
        print(f"   ✅ Seeded {len(STARTER_COMMENTS)} comment templates")
