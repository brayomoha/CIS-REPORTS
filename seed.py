"""
CIS School System - Database Seeder
=====================================
Run ONCE from the project root to set up:
  - All 12 grades (Reception → Grade 9)
  - RED and YELLOW streams for every grade
  - Academic Year 2026, Term 1, and 3 assessments
  - Default admin and principal accounts

Usage:
    python seed.py
"""

from werkzeug.security import generate_password_hash
from datetime import date
from app import create_app
from app.models import db, AcademicYear, Term, Assessment, Grade, Stream, Teacher


def seed():
    app = create_app()

    with app.app_context():
        print("\n🌱 Seeding CIS School System database...\n")

        # -------------------------------------------------------------------
        # ALL GRADES — in curriculum order
        # -------------------------------------------------------------------
        grade_definitions = [
            # ( name,       level_group,    sort_order )
            ("Reception", "reception",     1),
            ("PP1",        "preschool",    2),
            ("PP2",        "preschool",    3),
            ("Grade 1",    "lower_primary", 4),
            ("Grade 2",    "lower_primary", 5),
            ("Grade 3",    "lower_primary", 6),
            ("Grade 4",    "upper_primary", 7),
            ("Grade 5",    "upper_primary", 8),
            ("Grade 6",    "upper_primary", 9),
            ("Grade 7",    "junior",        10),
            ("Grade 8",    "junior",        11),
            ("Grade 9",    "junior",        12),
        ]

        for name, level_group, sort_order in grade_definitions:
            existing = Grade.query.filter_by(name=name).first()
            if existing:
                print(f"   ⏭️  {name} already exists — skipping")
                continue

            grade = Grade(name=name, level_group=level_group, sort_order=sort_order)
            db.session.add(grade)
            db.session.flush()

            for stream_name in ["RED", "YELLOW"]:
                db.session.add(Stream(grade_id=grade.id, name=stream_name))

            print(f"   ✅ {name:12s} ({level_group:15s})  →  RED + YELLOW streams")

        # -------------------------------------------------------------------
        # ACADEMIC YEAR 2026
        # -------------------------------------------------------------------
        year = AcademicYear.query.filter_by(year=2026).first()
        if not year:
            year = AcademicYear(year=2026, is_active=True)
            db.session.add(year)
            db.session.flush()
            print("\n   ✅ Created Academic Year 2026")
        else:
            print("\n   ⏭️  Academic Year 2026 already exists")

        # -------------------------------------------------------------------
        # TERM 1 — 2026
        # -------------------------------------------------------------------
        term1 = Term.query.filter_by(academic_year_id=year.id, term_number=1).first()
        if not term1:
            term1 = Term(
                academic_year_id=year.id,
                term_number=1,
                is_active=True,
                open_date=date(2026, 1, 6),
                close_date=date(2026, 3, 31),
                next_term_date=date(2026, 5, 4),
            )
            db.session.add(term1)
            db.session.flush()

            # Three assessments — Entry, Mid Term, End Term
            # Note: Reception only uses Mid Term (2) and End Term (3) in practice,
            # but all three are created. The mark-entry page filters appropriately.
            for num, name, open_default in [
                (1, "Entry Assessment", True),
                (2, "Mid Term",         True),
                (3, "End Term",         True),
            ]:
                db.session.add(Assessment(
                    term_id=term1.id,
                    name=name,
                    number=num,
                    is_open=open_default,
                ))
            print("   ✅ Created Term 1 — 2026 with 3 assessments")
        else:
            print("   ⏭️  Term 1 — 2026 already exists")

        # -------------------------------------------------------------------
        # DEFAULT ACCOUNTS
        # -------------------------------------------------------------------
        accounts = [
            ("CIS Administrator", "admin@cis.ac.ke",      "admin123",      "admin"),
            ("CIS Principal",     "principal@cis.ac.ke",  "principal123",  "principal"),
        ]
        print()
        for full_name, email, password, role in accounts:
            if Teacher.query.filter_by(email=email).first():
                print(f"   ⏭️  {email} already exists")
                continue
            db.session.add(Teacher(
                full_name=full_name,
                email=email,
                password_hash=generate_password_hash(password, method="pbkdf2:sha256"),
                role=role,
            ))
            print(f"   ✅ Created {role}: {email} / {password}")

        db.session.commit()

        print("\n" + "=" * 55)
        print("  ✅  Database ready!")
        print()
        print("  Grades created:  Reception, PP1, PP2,")
        print("                   Grade 1–6, Grade 7–9")
        print("  Streams:         RED + YELLOW for every grade")
        print()
        print("  Default logins:")
        print("    admin@cis.ac.ke          /  admin123")
        print("    principal@cis.ac.ke      /  principal123")
        print()
        print("  ⚠️  Change passwords after first login!")
        print("=" * 55 + "\n")


if __name__ == "__main__":
    seed()
