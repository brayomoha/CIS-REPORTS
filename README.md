# CIS School Management System
## Creative Integrated School — Nairobi, Kenya

---

## QUICK START (Local Mac)

```bash
cd ~/Downloads/cis_school_system
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
mkdir -p data uploads reports
python3 seed.py
python3 import_data.py
python3 import_end_term.py
python3 run.py
```

Open browser: http://127.0.0.1:5000

**Login:** admin@cis.ac.ke / admin123

---

## WHAT'S IN THIS SYSTEM

| Feature | Description |
|---|---|
| 12 Grade Levels | Reception, PP1, PP2, Grade 1–9 |
| 441 Students | All imported from your Excel files |
| 428 Comments | All teacher comments loaded |
| 434 End Term Marks | All imported |
| PDF Report Cards | Logo, correct colors, all 3 assessments |
| Download ALL PDFs | One ZIP for all 441 learners |
| Export to Excel | All marks backed up to Excel |
| Term Dates Editor | Closing date, next term date on PDFs |
| Stream Management | Add/rename/delete streams |
| Student Name Editing | Fix spelling errors |
| Teacher Accounts | Each teacher sees only their class |

---

## FILE STRUCTURE

```
cis_school_system/
├── run.py              ← Start server (local)
├── seed.py             ← Set up database (run once)
├── import_data.py      ← Load students + comments
├── import_end_term.py  ← Load end term marks
├── requirements.txt    ← Python libraries
├── Procfile            ← For Render.com deployment
│
├── app/
│   ├── models.py       ← Database structure
│   ├── grading.py      ← Grade calculations
│   ├── pdf_generator.py← PDF report cards
│   ├── comments_bank.py← Pre-written comments
│   │
│   ├── routes/
│   │   ├── admin.py    ← Manage teachers, students, streams, dates
│   │   ├── marks.py    ← Enter marks
│   │   ├── reports.py  ← Generate PDFs, download
│   │   ├── reception.py← Reception skill ratings
│   │   └── auth.py     ← Login/logout
│   │
│   ├── templates/      ← All HTML pages
│   └── static/img/     ← CIS logo
│
└── data/               ← Database (cis_school.db) — NEVER DELETE
```

---

## DEFAULT ACCOUNTS

| Role | Email | Password |
|---|---|---|
| Admin | admin@cis.ac.ke | admin123 |
| Principal | principal@cis.ac.ke | principal123 |

⚠️ Change passwords after first login!

---

## DEPLOYING ONLINE (Free — Render.com)

1. Create free account at render.com
2. Push this folder to GitHub
3. Connect GitHub repo to Render
4. Add environment variable: SECRET_KEY = any-random-string
5. Deploy — teachers access via your Render URL

