"""
CIS School System - Database Setup Script
Runs on Render startup. Safe to run multiple times.
"""
import os
import sys

# Fix postgres:// -> postgresql:// for SQLAlchemy
db_url = os.environ.get("DATABASE_URL", "")
if db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql://", 1)
    os.environ["DATABASE_URL"] = db_url

# If DATABASE_URL is set but looks wrong, warn and continue
if db_url and not db_url.startswith(("postgresql://", "sqlite://")):
    print(f"WARNING: Unusual DATABASE_URL format, attempting anyway...")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from app import create_app
    from app.models import db, Grade, Stream, AcademicYear, Term, Assessment, Teacher
    from werkzeug.security import generate_password_hash
    from datetime import date
except Exception as e:
    print(f"Import error: {e}")
    sys.exit(0)  # Don't crash - let gunicorn start anyway


def setup():
    try:
        app = create_app()
    except Exception as e:
        print(f"Could not create app: {e}")
        return

    with app.app_context():
        try:
            db.create_all()
            print("✅ Tables ready")
        except Exception as e:
            print(f"Could not create tables: {e}")
            return

        grade_defs = [
            ("Reception", "reception",     1),
            ("PP1",        "preschool",     2),
            ("PP2",        "preschool",     3),
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

        for name, level, order in grade_defs:
            try:
                grade = Grade.query.filter_by(name=name).first()
                if not grade:
                    grade = Grade(name=name, level_group=level, sort_order=order)
                    db.session.add(grade)
                    db.session.flush()
                    db.session.add(Stream(grade_id=grade.id, name="RED"))
                    db.session.add(Stream(grade_id=grade.id, name="YELLOW"))
                    print(f"   ✅ {name}")
            except Exception as e:
                print(f"   ⚠ {name}: {e}")
                db.session.rollback()

        try:
            yr = AcademicYear.query.filter_by(year=2026).first()
            if not yr:
                yr = AcademicYear(year=2026, is_active=True)
                db.session.add(yr)
                db.session.flush()

            term = Term.query.filter_by(academic_year_id=yr.id, term_number=1).first()
            if not term:
                term = Term(
                    academic_year_id=yr.id, term_number=1, is_active=True,
                    open_date=date(2026, 1, 6),
                    close_date=date(2026, 3, 31),
                    next_term_date=date(2026, 5, 4),
                )
                db.session.add(term)
                db.session.flush()
                for num, aname in [(1,"Entry Assessment"),(2,"Mid Term"),(3,"End Term")]:
                    db.session.add(Assessment(
                        term_id=term.id, name=aname, number=num, is_open=True))
        except Exception as e:
            print(f"   ⚠ Term setup: {e}")
            db.session.rollback()

        for full_name, email, password, role in [
            ("CIS Administrator", "admin@cis.ac.ke",     "admin123",     "admin"),
            ("CIS Principal",     "principal@cis.ac.ke", "principal123", "principal"),
        ]:
            try:
                if not Teacher.query.filter_by(email=email).first():
                    db.session.add(Teacher(
                        full_name=full_name, email=email,
                        password_hash=generate_password_hash(
                            password, method="pbkdf2:sha256"),
                        role=role,
                    ))
            except Exception as e:
                print(f"   ⚠ {email}: {e}")
                db.session.rollback()

        try:
            db.session.commit()
            print("✅ Database setup complete!")
            print("   admin@cis.ac.ke / admin123")
        except Exception as e:
            print(f"⚠ Final commit: {e}")
            db.session.rollback()


if __name__ == "__main__":
    setup()
