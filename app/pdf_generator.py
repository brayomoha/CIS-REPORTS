"""
CIS School System - PDF Report Card Generator
===============================================
Generates PDFs that exactly match the uploaded CIS Word templates:

HEADER (all levels):
  CREATIVE INTEGRATED SCHOOL
  ★★★ Kindergarten ★★★ Primary ★★★ Junior school
  ASSESSMENT Report
  P.O.BOX 1910-00900, KIAMBU-KENYA. TEL:0707044103/0720790555
  www.creativeintegratedschool.sc.ke

LEARNER INFO ROW:
  Learner's Name: ___  Grade/Class: ___  Term: ___
  Stream: ___  Teacher's Name: ___

MARKS TABLE — columns differ by level:
  PP1/PP2/Grade1-3:  LEARNING AREAS | ENTRY ASSESSMENT | MID TERM | END TERM
                     (each assessment = SCORE + LEVEL)
  Upper Primary:     same but with split Eng/Kisw
  Junior:            same with 8-band grading

GRADING KEY TABLE — exact bands from templates

TEACHER'S FEEDBACK:
  1. Learner's performance
  2. Learner's acquisition of core-competencies
  3. Learner's acquisition of values

DATES ROW:
  Next term commences on: ___    Closing date: ___

SIGNATURE LINE:
  Class Teacher Signature    Principal Signature    Parent's Signature
"""

import os
from datetime import date
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether, Image as RLImage
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.colors import HexColor, black, white
from flask import current_app
from .models import get_grade_level

LOGO_PATH = os.path.join(os.path.dirname(__file__), "static", "img", "cis_logo.png")

# ── EXACT COLOURS FROM CIS TEMPLATES ────────────────────────────────────────
NAVY   = HexColor("#002147")
GOLD   = HexColor("#C8962D")
WHITE  = white
BLACK  = black
LGREY  = HexColor("#F2F2F2")   # light row alternation
DGREY  = HexColor("#555555")   # muted text

# Level badge colours
LEVEL_CLR = {
    "EE":  HexColor("#1a6b1a"), "EE1": HexColor("#1a6b1a"), "EE2": HexColor("#2e8b2e"),
    "ME":  HexColor("#154f99"), "ME1": HexColor("#154f99"), "ME2": HexColor("#2968c8"),
    "AE":  HexColor("#9e5c00"), "AE1": HexColor("#9e5c00"), "AE2": HexColor("#c87800"),
    "BE":  HexColor("#990000"), "BE1": HexColor("#990000"), "BE2": HexColor("#c41010"),
    "NMP": HexColor("#7a0000"),
    "—":   DGREY,
}

W, H = A4   # 595 x 842 pts


# ── STYLE HELPERS ────────────────────────────────────────────────────────────

def S(name, **kw):
    """Quick ParagraphStyle factory."""
    return ParagraphStyle(name, **kw)

def bold(text, size=9, color=BLACK, align=TA_LEFT):
    return Paragraph(f"<b>{text}</b>", S("b", fontName="Helvetica-Bold",
                     fontSize=size, textColor=color, alignment=align))

def normal(text, size=8.5, color=BLACK, align=TA_LEFT):
    return Paragraph(text, S("n", fontName="Helvetica",
                     fontSize=size, textColor=color, alignment=align, leading=12))

def center(text, size=8.5, color=BLACK, bold_=False):
    fn = "Helvetica-Bold" if bold_ else "Helvetica"
    return Paragraph(text, S("c", fontName=fn, fontSize=size,
                     textColor=color, alignment=TA_CENTER, leading=11))

def level_cell(code):
    """Coloured performance level badge."""
    if not code or code == "—":
        return center("—", size=7.5, color=DGREY)
    clr = LEVEL_CLR.get(code, DGREY)
    return Paragraph(
        f'<font color="#{clr.hexval()[2:]}" ><b>{code}</b></font>',
        S("lv", fontName="Helvetica-Bold", fontSize=7.5,
          alignment=TA_CENTER, leading=10)
    )

def score_cell(val):
    if val is None or val == "":
        return center("—", size=8, color=DGREY)
    return center(str(int(round(float(val)))) if isinstance(val, float) else str(val),
                  size=8, bold_=True)


# ── HEADER — identical across all templates ──────────────────────────────────

def build_header(story):
    """
    Header matching the CIS Word template exactly:
    Logo (left) | School name + tagline + report title (centre) | empty (right)
    Contact details centred below.
    """
    LOGO_H = 2.6 * cm   # logo height on page
    LOGO_W = 2.6 * cm

    # Centre text column
    RED_CLR = HexColor("#CC0000")

    centre_text = [
        Paragraph("<b>CREATIVE INTEGRATED SCHOOL</b>",
                  S("sn", fontName="Helvetica-Bold", fontSize=16,
                    textColor=NAVY, alignment=TA_CENTER, leading=20, spaceAfter=4)),
        Paragraph(
            '<b><font color="#002147">Kindergarten </font>'
            '<font color="#CC0000">*Primary*</font>'
            '<font color="#002147"> Junior school</font></b>',
            S("st", fontName="Helvetica-Bold", fontSize=10,
              textColor=NAVY, alignment=TA_CENTER, leading=14, spaceAfter=4)),
        Paragraph("<b>ASSESSMENT Report</b>",
                  S("rt", fontName="Helvetica-Bold", fontSize=13,
                    textColor=NAVY, alignment=TA_CENTER, leading=16, spaceAfter=0)),
    ]

    # Logo image — use a blank spacer if file missing
    if os.path.exists(LOGO_PATH):
        logo = RLImage(LOGO_PATH, width=LOGO_W, height=LOGO_H)
    else:
        logo = Spacer(LOGO_W, LOGO_H)

    # Three-column header row: logo | text | mirror spacer for balance
    usable = W - 3.6 * cm   # approximate usable width
    side_w = LOGO_W + 0.4 * cm
    mid_w  = usable - side_w * 2

    header_row = Table(
        [[logo, centre_text, Spacer(side_w, LOGO_H)]],
        colWidths=[side_w, mid_w, side_w]
    )
    header_row.setStyle(TableStyle([
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN",         (1,0), (1,0),   "CENTER"),
        ("LEFTPADDING",   (0,0), (-1,-1), 0),
        ("RIGHTPADDING",  (0,0), (-1,-1), 0),
        ("TOPPADDING",    (0,0), (-1,-1), 0),
        ("BOTTOMPADDING", (0,0), (-1,-1), 0),
    ]))
    story.append(header_row)
    story.append(Spacer(1, 3))

    # Contact line centred below
    story.append(Paragraph(
        "<font color='#CC0000'><b>TEL:0707044103/0720790555</b></font>",
        S("ct", fontName="Helvetica-Bold", fontSize=10,
          textColor=HexColor("#CC0000"), alignment=TA_CENTER, spaceAfter=2)
    ))
    story.append(Paragraph(
        "P.O.BOX 1910-00900, KIAMBU-KENYA  |  www.creativeintegratedschool.sc.ke",
        S("ct2", fontName="Helvetica", fontSize=9,
          textColor=NAVY, alignment=TA_CENTER, spaceAfter=4)
    ))
    story.append(HRFlowable(width="100%", thickness=3,
                             color=HexColor("#CC0000"), spaceAfter=5))


# ── LEARNER INFO ROW ─────────────────────────────────────────────────────────

def build_info_row(story, report_data, usable_w):
    """Two-row learner info exactly matching the template layout."""
    name       = report_data["student_name"]
    grade      = report_data["grade"]
    stream     = report_data.get("stream", "")
    term_num   = report_data["term"]
    year       = report_data["year"]
    teacher    = report_data.get("class_teacher", "")
    adm        = report_data.get("admission_no", "")
    level      = get_grade_level(grade)

    # Row 1: Name | Grade/Class | Term
    grade_label = "Class" if level == "reception" else "Grade"
    grade_val   = f"{grade}  {stream}".strip() if stream else grade

    r1 = [[
        normal(f"<b>Learner's Name:</b>  {name}", size=8.5),
        normal(f"<b>{grade_label}:</b>  {grade_val}", size=8.5),
        normal(f"<b>Term:</b>  Term {term_num}  —  {year}", size=8.5),
    ]]
    t1 = Table(r1, colWidths=[usable_w * 0.48, usable_w * 0.27, usable_w * 0.25])
    t1.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ]))
    story.append(t1)

    # Row 2: Stream + Teacher (or Assessment No + Teacher for junior)
    if level == "junior":
        left_label = f"<b>Assessment No.:</b>  {adm}"
    else:
        left_label = f"<b>Stream:</b>  {stream}"

    r2 = [[
        normal(left_label, size=8.5),
        normal(f"<b>Teacher's Name:</b>  {teacher}", size=8.5),
        normal("", size=8.5),
    ]]
    t2 = Table(r2, colWidths=[usable_w * 0.38, usable_w * 0.37, usable_w * 0.25])
    t2.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 1),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ]))
    story.append(t2)


# ── GRADING KEY TABLE ────────────────────────────────────────────────────────

def build_key_table(level, usable_w):
    """Exact grading key from templates."""

    hdr_style = TableStyle([
        ("BACKGROUND",    (0,0), (-1,0), NAVY),
        ("TEXTCOLOR",     (0,0), (-1,0), WHITE),
        ("FONTNAME",      (0,0), (-1,-1), "Helvetica-Bold"),
        ("FONTSIZE",      (0,0), (-1,-1), 7.5),
        ("ALIGN",         (0,0), (-1,-1), "CENTER"),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
        ("GRID",          (0,0), (-1,-1), 0.4, HexColor("#cccccc")),
        ("TOPPADDING",    (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
    ])

    if level == "reception":
        data = [
            ["Exceeding Expectation", "Meeting Expectation",
             "Approaching Expectation", "Needs More Practice"],
            ["4", "3", "2", "1"],
        ]
        cw = [usable_w / 4] * 4

    elif level == "junior":
        data = [
            ["EE1\n90–100", "EE2\n75–89", "ME1\n58–74", "ME2\n41–57",
             "AE1\n31–40", "AE2\n21–30", "BE1\n11–20", "BE2\n0–10"],
            ["Exceeds", "Exceeds", "Meets", "Meets",
             "Approaches", "Approaches", "Below", "Below"],
        ]
        cw = [usable_w / 8] * 8

    else:
        # 4-band — exact from templates
        data = [
            ["KEY", "EE (4)", "ME (3)", "AE (2)", "BE (1)"],
            ["", "Exceeding Expectation", "Meeting Expectation",
             "Approaching Expectation", "Below Expectation"],
            ["", "26–30", "18–25", "11–17", "0–10"],
        ]
        cw = [1.5*cm, usable_w*0.23, usable_w*0.23, usable_w*0.23, usable_w*0.23]
        # Colour the EE/ME/AE/BE headers
        hdr_style = TableStyle([
            ("BACKGROUND",    (0,0), (-1,0), NAVY),
            ("TEXTCOLOR",     (0,0), (-1,0), WHITE),
            ("BACKGROUND",    (0,0), (0,-1), NAVY),
            ("TEXTCOLOR",     (0,0), (0,-1), WHITE),
            ("FONTNAME",      (0,0), (-1,-1), "Helvetica-Bold"),
            ("FONTSIZE",      (0,0), (-1,-1), 7.5),
            ("ALIGN",         (0,0), (-1,-1), "CENTER"),
            ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
            ("GRID",          (0,0), (-1,-1), 0.4, HexColor("#cccccc")),
            ("ROWBACKGROUNDS",(0,1), (-1,-1), [LGREY, WHITE]),
            ("TOPPADDING",    (0,0), (-1,-1), 3),
            ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ])

    t = Table([[center(str(c), size=7.5, bold_=True) for c in row] for row in data],
              colWidths=cw)
    t.setStyle(hdr_style)
    return t


# ── MARKS TABLE ──────────────────────────────────────────────────────────────

def build_marks_table(report_data, usable_w):
    """
    Builds the subject marks table matching the exact column layout
    from each template level.
    """
    level      = get_grade_level(report_data["grade"])
    subjects   = report_data["subjects"]
    split      = report_data.get("split_subjects", {})
    assessments = report_data["assessments"]   # {1: summary, 2: summary, 3: summary}

    # All levels get all 3 assessments
    ass_cols = [(1, "ENTRY ASSESSMENT"), (2, "MID TERM"), (3, "END TERM")]

    num_ass = len(ass_cols)

    # ── Column widths ────────────────────────────────
    subj_w   = 5.0 * cm
    score_w  = 1.3 * cm
    level_w  = 1.3 * cm
    pair_w   = score_w + level_w

    total_pair = usable_w - subj_w
    pair_unit  = total_pair / num_ass
    score_w    = pair_unit * 0.52
    level_w    = pair_unit * 0.48

    col_widths = [subj_w] + [score_w, level_w] * num_ass

    # ── Header rows ──────────────────────────────────
    h1 = [bold("LEARNING AREAS", size=7.5, color=WHITE)]
    h2 = [bold("", size=7)]
    for _, label in ass_cols:
        h1 += [bold(label, size=7.5, color=WHITE, align=TA_CENTER), bold("", size=7)]
        h2 += [bold("SCORE", size=7, color=WHITE, align=TA_CENTER),
               bold("LEVEL", size=7, color=WHITE, align=TA_CENTER)]

    rows = [h1, h2]

    # ── Data rows ────────────────────────────────────
    for subj in subjects:
        is_split = subj in split
        if is_split:
            p1l, p1m, p2l, p2m = split[subj]
            # Show subject name with paper label only (no insha/composition mention)
            subj_cell = normal(
                f"{subj}",
                size=8
            )
        else:
            subj_cell = normal(subj, size=8)

        row = [subj_cell]
        for ass_num, _ in ass_cols:
            summary   = assessments.get(ass_num, {})
            subj_data = summary.get("subjects", {}).get(subj, {})
            code      = subj_data.get("grade_code", "—")

            # For split subjects (Eng/Kisw): show only the main paper score
            # not the insha/composition combined percentage
            if is_split:
                sc = subj_data.get("paper1")
            else:
                sc = subj_data.get("score")

            row += [score_cell(sc), level_cell(code)]

        rows.append(row)

    # ── Span header cells ────────────────────────────
    span_cmds = [("SPAN", (0,0), (0,1))]   # subject column spans both header rows
    for i in range(num_ass):
        col_start = 1 + i * 2
        span_cmds.append(("SPAN", (col_start, 0), (col_start+1, 0)))

    table = Table(rows, colWidths=col_widths, repeatRows=2)
    table.setStyle(TableStyle([
        # Header background
        ("BACKGROUND",    (0,0), (-1,1), NAVY),
        ("TEXTCOLOR",     (0,0), (-1,1), WHITE),
        # Data rows alternating
        ("ROWBACKGROUNDS",(0,2), (-1,-1), [WHITE, LGREY]),
        ("FONTNAME",      (0,2), (-1,-1), "Helvetica"),
        ("FONTSIZE",      (0,2), (-1,-1), 8),
        # Grid
        ("GRID",          (0,0), (-1,-1), 0.4, HexColor("#bbbbbb")),
        # Padding
        ("LEFTPADDING",   (0,0), (-1,-1), 4),
        ("RIGHTPADDING",  (0,0), (-1,-1), 4),
        ("TOPPADDING",    (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN",         (0,0), (0,-1), "LEFT"),
        # Span commands
        *span_cmds,
    ]))
    return table


# ── RECEPTION SKILLS TABLE ───────────────────────────────────────────────────

def build_reception_skills_table(report_data, usable_w):
    """Builds the skills rating table for Reception."""
    sections      = report_data.get("sections", {})
    skills_by_ass = report_data.get("skills_by_assessment", {})
    ass_cols      = [(2, "MID TERM"), (3, "END TERM")]

    RATING_LABEL = {4: "EE", 3: "ME", 2: "AE", 1: "NMP"}

    story_parts = []
    skill_w = 6.5 * cm
    col_w   = (usable_w - skill_w) / len(ass_cols)

    for section, items in sections.items():
        # Section header row
        header = [[
            bold(section.upper(), size=7.5, color=WHITE),
            *[bold(label, size=7.5, color=WHITE, align=TA_CENTER)
              for _, label in ass_cols]
        ]]
        data_rows = []
        for item in items:
            row = [normal(item, size=8)]
            for ass_num, _ in ass_cols:
                ratings = skills_by_ass.get(ass_num, {})
                r = ratings.get(item)
                code = r.code if r else "—"
                row.append(level_cell(code))
            data_rows.append(row)

        all_rows = header + data_rows
        t = Table(all_rows, colWidths=[skill_w] + [col_w]*len(ass_cols))
        t.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,0), NAVY),
            ("TEXTCOLOR",     (0,0), (-1,0), WHITE),
            ("ROWBACKGROUNDS",(0,1), (-1,-1), [WHITE, LGREY]),
            ("GRID",          (0,0), (-1,-1), 0.4, HexColor("#bbbbbb")),
            ("LEFTPADDING",   (0,0), (-1,-1), 4),
            ("RIGHTPADDING",  (0,0), (-1,-1), 4),
            ("TOPPADDING",    (0,0), (-1,-1), 3),
            ("BOTTOMPADDING", (0,0), (-1,-1), 3),
            ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
            ("ALIGN",         (0,0), (0,-1), "LEFT"),
        ]))
        story_parts.append(t)
        story_parts.append(Spacer(1, 4))

    return story_parts


# ── FEEDBACK SECTION ─────────────────────────────────────────────────────────

def build_feedback(story, comments, level):
    story.append(Paragraph(
        "<b>Teacher's Feedback on:</b>",
        S("fh", fontName="Helvetica-Bold", fontSize=9,
          textColor=NAVY, spaceBefore=6, spaceAfter=4)
    ))

    if level == "reception":
        story.append(normal("<b>General Comments:</b>", size=8.5))
        text = comments.get("general", "").strip()
        story.append(normal(text or " ", size=8.5))
        story.append(Spacer(1, 6))
    else:
        sections = [
            ("1. Learner's performance",                       comments.get("performance", "")),
            ("2. Learner's acquisition of core-competencies",  comments.get("competencies", "")),
            ("3. Learner's acquisition of values",             comments.get("values", "")),
        ]
        for heading, text in sections:
            story.append(normal(f"<b>{heading}</b>", size=8.5))
            story.append(normal(text.strip() if text else " ", size=8.5))
            story.append(Spacer(1, 5))


# ── DATES ROW ────────────────────────────────────────────────────────────────

def build_dates(story, report_data, usable_w):
    close_date     = report_data.get("close_date")
    next_term_date = report_data.get("next_term_date")

    def fmt(d):
        if not d:
            return "…………………………"
        if isinstance(d, date):
            # e.g. "Tuesday 31st March, 2026"
            day = d.day
            suffix = "th" if 11 <= day <= 13 else {1:"st",2:"nd",3:"rd"}.get(day%10,"th")
            return d.strftime(f"%-d{suffix} %B, %Y")
        return str(d)

    row = [[
        normal(f"<b>Next term commences on:</b>  {fmt(next_term_date)}", size=8),
        normal(f"<b>Closing date:</b>  {fmt(close_date)}", size=8),
    ]]
    t = Table(row, colWidths=[usable_w * 0.55, usable_w * 0.45])
    t.setStyle(TableStyle([
        ("LEFTPADDING",   (0,0), (-1,-1), 0),
        ("RIGHTPADDING",  (0,0), (-1,-1), 0),
        ("TOPPADDING",    (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
    ]))
    story.append(t)
    story.append(Spacer(1, 10))


# ── SIGNATURE LINE ───────────────────────────────────────────────────────────

def build_signatures(story, usable_w):
    sig = Table([[
        normal("______________________", size=8, align=TA_CENTER),
        normal("_________________________", size=8, align=TA_CENTER),
        normal("___________________", size=8, align=TA_CENTER),
    ]], colWidths=[usable_w/3]*3)
    lbl = Table([[
        normal("Class Teacher Signature", size=7.5, color=DGREY, align=TA_CENTER),
        normal("Principal Signature", size=7.5, color=DGREY, align=TA_CENTER),
        normal("Parent's Signature", size=7.5, color=DGREY, align=TA_CENTER),
    ]], colWidths=[usable_w/3]*3)
    for t in [sig, lbl]:
        t.setStyle(TableStyle([
            ("ALIGN",   (0,0), (-1,-1), "CENTER"),
            ("VALIGN",  (0,0), (-1,-1), "MIDDLE"),
            ("TOPPADDING",    (0,0), (-1,-1), 3),
            ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ]))
    story.append(sig)
    story.append(lbl)


# ── MAIN ENTRY POINT ─────────────────────────────────────────────────────────

def generate_report_pdf(report_data, student):
    """
    Generate one PDF report card matching the exact CIS template.
    Returns the file path.
    """
    reports_folder = current_app.config["REPORTS_FOLDER"]
    os.makedirs(reports_folder, exist_ok=True)

    safe_name = student.full_name.replace(" ", "_").replace("/", "-")
    filename  = f"{student.admission_no}_{safe_name}_T{report_data['term']}_{report_data['year']}.pdf"
    filepath  = os.path.join(reports_folder, filename)

    L_MARGIN = R_MARGIN = 1.8 * cm
    T_MARGIN = B_MARGIN = 1.5 * cm
    usable_w = W - L_MARGIN - R_MARGIN

    doc = SimpleDocTemplate(
        filepath, pagesize=A4,
        leftMargin=L_MARGIN, rightMargin=R_MARGIN,
        topMargin=T_MARGIN, bottomMargin=B_MARGIN,
    )

    level   = get_grade_level(report_data["grade"])
    story   = []
    comments = report_data.get("comments", {})

    # ── HEADER ──────────────────────────────────────
    build_header(story)

    # ── LEARNER INFO ─────────────────────────────────
    build_info_row(story, report_data, usable_w)

    # ── GRADING KEY (before or after table — matches template order) ──
    if level != "reception":
        story.append(build_key_table(level, usable_w))
        story.append(Spacer(1, 6))

    # ── MARKS / SKILLS TABLE ─────────────────────────
    if level == "reception":
        story.append(build_key_table(level, usable_w))
        story.append(Spacer(1, 6))
        for part in build_reception_skills_table(report_data, usable_w):
            story.append(part)
    else:
        story.append(build_marks_table(report_data, usable_w))

    story.append(Spacer(1, 8))

    # ── TEACHER FEEDBACK ────────────────────────────
    build_feedback(story, comments, level)

    # ── DATES ────────────────────────────────────────
    build_dates(story, report_data, usable_w)

    # ── SIGNATURES ───────────────────────────────────
    build_signatures(story, usable_w)

    doc.build(story)
    return filepath
