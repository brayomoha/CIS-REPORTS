"""
CIS School System - App Factory
=================================
Creates and configures the Flask web application.
"""

import os
from flask import Flask
from .models import db


def create_app(config=None):
    app = Flask(__name__, template_folder="templates", static_folder="static")

    # -----------------------------------------------------------------------
    # CONFIGURATION
    # -----------------------------------------------------------------------
    app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "cis-dev-secret-change-in-production")

    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    db_path  = os.path.join(base_dir, "data", "cis_school.db")

    # Build database URL - handle all Render PostgreSQL URL formats
    database_url = f"sqlite:///{db_path}"  # default fallback
    raw_url = os.environ.get("DATABASE_URL", "")
    if raw_url:
        # Fix postgres:// -> postgresql:// (Render uses old format)
        if raw_url.startswith("postgres://"):
            raw_url = raw_url.replace("postgres://", "postgresql://", 1)
        # Only use if it looks valid
        if raw_url.startswith(("postgresql://", "sqlite://")):
            database_url = raw_url
        else:
            print(f"WARNING: Ignoring invalid DATABASE_URL, using SQLite")
    app.config["SQLALCHEMY_DATABASE_URI"] = database_url
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
    app.config["UPLOAD_FOLDER"]                  = os.path.join(base_dir, "uploads")
    app.config["REPORTS_FOLDER"]                 = os.path.join(base_dir, "reports")
    app.config["MAX_CONTENT_LENGTH"]             = 16 * 1024 * 1024   # 16 MB

    # Create required folders if they don't exist (important for cloud deployment)
    for folder in [os.path.join(base_dir, "data"),
                   os.path.join(base_dir, "uploads"),
                   os.path.join(base_dir, "reports")]:
        os.makedirs(folder, exist_ok=True)

    if config:
        app.config.update(config)

    # -----------------------------------------------------------------------
    # EXTENSIONS
    # -----------------------------------------------------------------------
    db.init_app(app)

    # -----------------------------------------------------------------------
    # BLUEPRINTS
    # -----------------------------------------------------------------------
    from .routes.auth      import auth_bp
    from .routes.admin     import admin_bp
    from .routes.marks     import marks_bp
    from .routes.reports   import reports_bp
    from .routes.main      import main_bp
    from .routes.reception import reception_bp
    from .routes.upload    import upload_bp
    from .routes.setup     import setup_bp

    app.register_blueprint(main_bp)
    app.register_blueprint(auth_bp,       url_prefix="/auth")
    app.register_blueprint(admin_bp,      url_prefix="/admin")
    app.register_blueprint(marks_bp,      url_prefix="/marks")
    app.register_blueprint(reports_bp,    url_prefix="/reports")
    app.register_blueprint(reception_bp,  url_prefix="/reception")
    app.register_blueprint(upload_bp,     url_prefix="/upload")
    app.register_blueprint(setup_bp,      url_prefix="/setup")

    # -----------------------------------------------------------------------
    # CREATE TABLES + SEED COMMENT TEMPLATES
    # -----------------------------------------------------------------------
    with app.app_context():
        # Import CommentTemplate here so the table is registered
        from .comments_bank import CommentTemplate, seed_comment_templates
        db.create_all()
        seed_comment_templates(app)

    # Auto-create tables and seed on first run
    with app.app_context():
        try:
            db.create_all()
            _migrate_db(app)
            _seed_if_empty(app)
        except Exception as e:
            print(f"DB init warning: {e}")

    return app


def _migrate_db(app):
    """Add any missing columns. Safe to run multiple times."""
    from sqlalchemy import text, inspect as sa_inspect
    with app.app_context():
        try:
            db.create_all()
            inspector = sa_inspect(db.engine)
            if "teachers" in inspector.get_table_names():
                cols = [c["name"] for c in inspector.get_columns("teachers")]
                if "subject" not in cols:
                    with db.engine.connect() as conn:
                        conn.execute(text(
                            "ALTER TABLE teachers ADD COLUMN subject VARCHAR(200)"
                        ))
                        conn.commit()
                    print("✅ Added subject column to teachers")
                if "extra_streams" not in cols:
                    with db.engine.connect() as conn:
                        conn.execute(text(
                            "ALTER TABLE teachers ADD COLUMN extra_streams VARCHAR(200)"
                        ))
                        conn.commit()
                    print("✅ Added extra_streams column to teachers")
        except Exception as e:
            print(f"Migration note: {e}")


def _seed_if_empty(app):
    """Seed grades, term and admin accounts if database is empty."""
    from .models import Grade, Stream, AcademicYear, Term, Assessment, Teacher
    from werkzeug.security import generate_password_hash
    from datetime import date

    try:
        # If grades already exist, skip
        if Grade.query.count() > 0:
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
            g = Grade(name=name, level_group=level, sort_order=order)
            db.session.add(g)
            db.session.flush()
            db.session.add(Stream(grade_id=g.id, name="RED"))
            db.session.add(Stream(grade_id=g.id, name="YELLOW"))

        yr = AcademicYear(year=2026, is_active=True)
        db.session.add(yr)
        db.session.flush()

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

        for full_name, email, password, role in [
            ("CIS Administrator", "admin@cis.ac.ke",     "admin123",     "admin"),
            ("CIS Principal",     "principal@cis.ac.ke", "principal123", "principal"),
        ]:
            db.session.add(Teacher(
                full_name=full_name, email=email,
                password_hash=generate_password_hash(password, method="pbkdf2:sha256"),
                role=role,
            ))

        db.session.commit()
        print("✅ Database seeded successfully")

    except Exception as e:
        db.session.rollback()
        print(f"Seed warning: {e}")
