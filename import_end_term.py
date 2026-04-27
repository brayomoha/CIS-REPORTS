"""
CIS School System - Import End Term Marks
==========================================
Run this ONCE to load all End Term 1 marks into the database.

Usage:
    python3 import_end_term.py
"""

from app import create_app
from app.models import db, Grade, Stream, Student, Term, Assessment, Mark
from app.grading import assign_performance_level, combine_split_subject

END_TERM_DATA = {
  "Grade 9": [
    {
      "name": "ABEL CHERUIYOT",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 40.0,
          "p2": 15.0,
          "pct": 55.0
        },
        "English": {
          "p1": 42.0,
          "p2": 35.0,
          "pct": 77.0
        },
        "Mathematics": 26.0,
        "Social Studies": 79.0,
        "Christian Religious Education": 85.0,
        "Creative Arts": 77.0,
        "Integrated Science": 69.0,
        "Agriculture": 77.0,
        "Pre-Technical Studies": 75.0
      }
    },
    {
      "name": "BENEDICT WANGAI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 38.0,
          "p2": 25.0,
          "pct": 63.0
        },
        "English": {
          "p1": 46.0,
          "p2": 39.0,
          "pct": 85.0
        },
        "Mathematics": 92.0,
        "Social Studies": 94.0,
        "Christian Religious Education": 91.0,
        "Creative Arts": 93.0,
        "Integrated Science": 88.0,
        "Agriculture": 89.0,
        "Pre-Technical Studies": 84.0
      }
    },
    {
      "name": "BRIGHTON AMBALE",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 35.0,
          "p2": 22.0,
          "pct": 57.0
        },
        "English": {
          "p1": 28.0,
          "p2": 26.0,
          "pct": 54.0
        },
        "Mathematics": 63.0,
        "Social Studies": 82.0,
        "Christian Religious Education": 66.0,
        "Creative Arts": 78.0,
        "Integrated Science": 62.0,
        "Agriculture": 67.0,
        "Pre-Technical Studies": 70.0
      }
    },
    {
      "name": "CATHERINE NYAKAIRO",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 37.0,
          "p2": 22.0,
          "pct": 59.0
        },
        "English": {
          "p1": 39.0,
          "p2": 35.0,
          "pct": 74.0
        },
        "Mathematics": 60.0,
        "Social Studies": 83.0,
        "Christian Religious Education": 67.0,
        "Creative Arts": 73.0,
        "Integrated Science": 95.0,
        "Agriculture": 86.0,
        "Pre-Technical Studies": 76.0
      }
    },
    {
      "name": "CHRISTIAN WANGILA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 36.0,
          "p2": 20.0,
          "pct": 56.0
        },
        "English": {
          "p1": 37.0,
          "p2": 31.0,
          "pct": 68.0
        },
        "Mathematics": 74.0,
        "Social Studies": 85.0,
        "Christian Religious Education": 72.0,
        "Creative Arts": 79.0,
        "Integrated Science": 63.0,
        "Agriculture": 66.0,
        "Pre-Technical Studies": 76.0
      }
    },
    {
      "name": "DAMARIS KAJEMBA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 38.0,
          "p2": 26.0,
          "pct": 64.0
        },
        "English": {
          "p1": 36.0,
          "p2": 29.0,
          "pct": 65.0
        },
        "Mathematics": 37.0,
        "Social Studies": 73.0,
        "Christian Religious Education": 82.0,
        "Creative Arts": 70.0,
        "Integrated Science": 72.0,
        "Agriculture": 70.0,
        "Pre-Technical Studies": 74.0
      }
    },
    {
      "name": "EMMANUEL HANS",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 35.0,
          "p2": 29.0,
          "pct": 64.0
        },
        "English": {
          "p1": 41.0,
          "p2": 38.0,
          "pct": 79.0
        },
        "Mathematics": 73.0,
        "Social Studies": 87.0,
        "Christian Religious Education": 78.0,
        "Creative Arts": 72.0,
        "Integrated Science": 81.0,
        "Agriculture": 71.0,
        "Pre-Technical Studies": 85.0
      }
    },
    {
      "name": "GABRIEL MUKOPI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 42.0,
          "p2": 37.0,
          "pct": 79.0
        },
        "English": {
          "p1": 49.0,
          "p2": 42.0,
          "pct": 91.0
        },
        "Mathematics": 96.0,
        "Social Studies": 95.0,
        "Christian Religious Education": 97.0,
        "Creative Arts": 93.0,
        "Integrated Science": 99.0,
        "Agriculture": 83.0,
        "Pre-Technical Studies": 94.0
      }
    },
    {
      "name": "GABRIELLA WANJIRU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 36.0,
          "p2": 17.0,
          "pct": 53.0
        },
        "English": {
          "p1": 40.0,
          "p2": 36.0,
          "pct": 76.0
        },
        "Mathematics": 21.0,
        "Social Studies": 71.0,
        "Christian Religious Education": 72.0,
        "Creative Arts": 66.0,
        "Integrated Science": 68.0,
        "Agriculture": 69.0,
        "Pre-Technical Studies": 75.0
      }
    },
    {
      "name": "HAILEY WANJIKU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 22.0,
          "p2": 10.0,
          "pct": 32.0
        },
        "English": {
          "p1": 33.0,
          "p2": 30.0,
          "pct": 63.0
        },
        "Mathematics": 31.0,
        "Social Studies": 58.0,
        "Christian Religious Education": 61.0,
        "Creative Arts": 58.0,
        "Integrated Science": 66.0,
        "Agriculture": 66.0,
        "Pre-Technical Studies": 64.0
      }
    },
    {
      "name": "HAROLD REMMY",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 45.0,
          "p2": 25.0,
          "pct": 70.0
        },
        "English": {
          "p1": 38.0,
          "p2": 34.0,
          "pct": 72.0
        },
        "Mathematics": 63.0,
        "Social Studies": 81.0,
        "Christian Religious Education": 85.0,
        "Creative Arts": 81.0,
        "Integrated Science": 86.0,
        "Agriculture": 84.0,
        "Pre-Technical Studies": 89.0
      }
    },
    {
      "name": "HELLEN WAITHERA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 35.0,
          "p2": 13.0,
          "pct": 48.0
        },
        "English": {
          "p1": 38.0,
          "p2": 31.0,
          "pct": 69.0
        },
        "Mathematics": 20.0,
        "Social Studies": 64.0,
        "Christian Religious Education": 63.0,
        "Creative Arts": 62.0,
        "Integrated Science": 61.0,
        "Agriculture": 72.0,
        "Pre-Technical Studies": 58.0
      }
    },
    {
      "name": "JAIDEN NJENGA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 34.0,
          "p2": 19.0,
          "pct": 53.0
        },
        "English": {
          "p1": 38.0,
          "p2": 37.0,
          "pct": 75.0
        },
        "Mathematics": 66.0,
        "Social Studies": 83.0,
        "Christian Religious Education": 84.0,
        "Creative Arts": 73.0,
        "Integrated Science": 78.0,
        "Agriculture": 67.0,
        "Pre-Technical Studies": 73.0
      }
    },
    {
      "name": "KEVIN MATALANGA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 27.0,
          "p2": 16.0,
          "pct": 43.0
        },
        "English": {
          "p1": 37.0,
          "p2": 30.0,
          "pct": 67.0
        },
        "Mathematics": 45.0,
        "Social Studies": 64.0,
        "Christian Religious Education": 63.0,
        "Creative Arts": 63.0,
        "Integrated Science": 64.0,
        "Agriculture": 61.0,
        "Pre-Technical Studies": 73.0
      }
    },
    {
      "name": "KEZZY EMMANUELLA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 38.0,
          "p2": 24.0,
          "pct": 62.0
        },
        "English": {
          "p1": 44.0,
          "p2": 41.0,
          "pct": 85.0
        },
        "Mathematics": 51.0,
        "Social Studies": 90.0,
        "Christian Religious Education": 86.0,
        "Creative Arts": 91.0,
        "Integrated Science": 71.0,
        "Agriculture": 81.0,
        "Pre-Technical Studies": 85.0
      }
    },
    {
      "name": "LEONE GICHERU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 31.0,
          "p2": 9.0,
          "pct": 40.0
        },
        "English": {
          "p1": 38.0,
          "p2": 31.0,
          "pct": 69.0
        },
        "Mathematics": 34.0,
        "Social Studies": 48.0,
        "Christian Religious Education": 68.0,
        "Creative Arts": 66.0,
        "Integrated Science": 58.0,
        "Agriculture": 79.0,
        "Pre-Technical Studies": 83.0
      }
    },
    {
      "name": "RAYNA HANANI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 38.0,
          "p2": 23.0,
          "pct": 61.0
        },
        "English": {
          "p1": 47.0,
          "p2": 40.0,
          "pct": 87.0
        },
        "Mathematics": 69.0,
        "Social Studies": 92.0,
        "Christian Religious Education": 92.0,
        "Creative Arts": 81.0,
        "Integrated Science": 89.0,
        "Agriculture": 77.0,
        "Pre-Technical Studies": 89.0
      }
    },
    {
      "name": "SHALOM GIKUBU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 27.0,
          "p2": 18.0,
          "pct": 45.0
        },
        "English": {
          "p1": 42.0,
          "p2": 42.0,
          "pct": 84.0
        },
        "Mathematics": 91.0,
        "Social Studies": 91.0,
        "Christian Religious Education": 81.0,
        "Creative Arts": 82.0,
        "Integrated Science": 84.0,
        "Agriculture": 87.0,
        "Pre-Technical Studies": 90.0
      }
    },
    {
      "name": "SHAMMAH CHEGE",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 29.0,
          "p2": 6.0,
          "pct": 35.0
        },
        "English": {
          "p1": 35.0,
          "p2": 21.0,
          "pct": 56.0
        },
        "Mathematics": 33.0,
        "Social Studies": 65.0,
        "Christian Religious Education": 64.0,
        "Creative Arts": 45.0,
        "Integrated Science": 50.0,
        "Agriculture": 61.0,
        "Pre-Technical Studies": 76.0
      }
    },
    {
      "name": "SHAUN MBURU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 26.0,
          "p2": 6.0,
          "pct": 32.0
        },
        "English": {
          "p1": 40.0,
          "p2": 35.0,
          "pct": 75.0
        },
        "Mathematics": 70.0,
        "Social Studies": 87.0,
        "Christian Religious Education": 76.0,
        "Creative Arts": 80.0,
        "Integrated Science": 80.0,
        "Agriculture": 77.0,
        "Pre-Technical Studies": 75.0
      }
    },
    {
      "name": "VICTORIA ODHIAMBO",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 41.0,
          "p2": 20.0,
          "pct": 61.0
        },
        "English": {
          "p1": 42.0,
          "p2": 37.0,
          "pct": 79.0
        },
        "Mathematics": 70.0,
        "Social Studies": 91.0,
        "Christian Religious Education": 88.0,
        "Creative Arts": 85.0,
        "Integrated Science": 91.0,
        "Agriculture": 79.0,
        "Pre-Technical Studies": 81.0
      }
    },
    {
      "name": "SUBJECT RANK",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 9.0
        },
        "English": {
          "p1": null,
          "p2": null,
          "pct": 7.0
        },
        "Mathematics": 8.0,
        "Social Studies": 1.0,
        "Christian Religious Education": 3.0,
        "Creative Arts": 6.0,
        "Integrated Science": 4.0,
        "Agriculture": 5.0,
        "Pre-Technical Studies": 2.0
      }
    }
  ],
  "Grade 8": [
    {
      "name": "ADDIE CYANN WANGUI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 23.0,
          "p2": 40.0,
          "pct": 63.0
        },
        "English": {
          "p1": 38.0,
          "p2": 33.0,
          "pct": 71.0
        },
        "Mathematics": 67.0,
        "Social Studies": 88.0,
        "Christian Religious Education": 79.0,
        "Creative Arts": 81.0,
        "Integrated Science": 77.0,
        "Agriculture": 79.0,
        "Pre-Technical Studies": 90.0
      }
    },
    {
      "name": "AGLA KAHOSI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 20.0,
          "p2": 41.0,
          "pct": 61.0
        },
        "English": {
          "p1": 30.0,
          "p2": 19.0,
          "pct": 49.0
        },
        "Mathematics": 37.0,
        "Social Studies": 47.0,
        "Christian Religious Education": null,
        "Creative Arts": 45.0,
        "Integrated Science": 39.0,
        "Agriculture": 55.0,
        "Pre-Technical Studies": 40.0
      }
    },
    {
      "name": "ANDREW JAYDEN MAINA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 29.0,
          "p2": 34.0,
          "pct": 63.0
        },
        "English": {
          "p1": 36.0,
          "p2": 27.0,
          "pct": 63.0
        },
        "Mathematics": 50.0,
        "Social Studies": 80.0,
        "Christian Religious Education": 51.0,
        "Creative Arts": 66.0,
        "Integrated Science": 63.0,
        "Agriculture": 59.0,
        "Pre-Technical Studies": 70.0
      }
    },
    {
      "name": "ANN NYAMBURA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 33.0,
          "p2": 46.0,
          "pct": 79.0
        },
        "English": {
          "p1": 42.0,
          "p2": 33.0,
          "pct": 75.0
        },
        "Mathematics": 70.0,
        "Social Studies": 76.0,
        "Christian Religious Education": 94.0,
        "Creative Arts": 87.0,
        "Integrated Science": 77.0,
        "Agriculture": 85.0,
        "Pre-Technical Studies": 88.0
      }
    },
    {
      "name": "ANTHONY RAPHA MUITO",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 21.0,
          "p2": 33.0,
          "pct": 54.0
        },
        "English": {
          "p1": 33.0,
          "p2": 27.0,
          "pct": 60.0
        },
        "Mathematics": 62.0,
        "Social Studies": 88.0,
        "Christian Religious Education": 44.0,
        "Creative Arts": 67.0,
        "Integrated Science": 60.0,
        "Agriculture": 68.0,
        "Pre-Technical Studies": 70.0
      }
    },
    {
      "name": "BRIAN MACHARIA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 20.0,
          "p2": 34.0,
          "pct": 54.0
        },
        "English": {
          "p1": 36.0,
          "p2": 34.0,
          "pct": 70.0
        },
        "Mathematics": 41.0,
        "Social Studies": 62.0,
        "Christian Religious Education": 77.0,
        "Creative Arts": 62.0,
        "Integrated Science": 51.0,
        "Agriculture": 71.0,
        "Pre-Technical Studies": 76.0
      }
    },
    {
      "name": "CINDY NJURA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 37.0,
          "p2": 41.0,
          "pct": 78.0
        },
        "English": {
          "p1": 42.0,
          "p2": 33.0,
          "pct": 75.0
        },
        "Mathematics": 84.0,
        "Social Studies": 89.0,
        "Christian Religious Education": 88.0,
        "Creative Arts": 77.0,
        "Integrated Science": 86.0,
        "Agriculture": 70.0,
        "Pre-Technical Studies": 89.0
      }
    },
    {
      "name": "ELYANA NJERI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 25.0,
          "p2": 41.0,
          "pct": 66.0
        },
        "English": {
          "p1": 37.0,
          "p2": 33.0,
          "pct": 70.0
        },
        "Mathematics": 80.0,
        "Social Studies": 91.0,
        "Christian Religious Education": 75.0,
        "Creative Arts": 79.0,
        "Integrated Science": 80.0,
        "Agriculture": 70.0,
        "Pre-Technical Studies": 89.0
      }
    },
    {
      "name": "EMMANUEL LUKUYU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 19.0,
          "p2": 39.0,
          "pct": 58.0
        },
        "English": {
          "p1": 29.0,
          "p2": 23.0,
          "pct": 52.0
        },
        "Mathematics": 34.0,
        "Social Studies": 63.0,
        "Christian Religious Education": 59.0,
        "Creative Arts": 57.0,
        "Integrated Science": 40.0,
        "Agriculture": 55.0,
        "Pre-Technical Studies": 54.0
      }
    },
    {
      "name": "ESTHER OCHIENG’",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 31.0,
          "p2": 45.0,
          "pct": 76.0
        },
        "English": {
          "p1": 28.0,
          "p2": 30.0,
          "pct": 58.0
        },
        "Mathematics": 33.0,
        "Social Studies": 68.0,
        "Christian Religious Education": 79.0,
        "Creative Arts": 63.0,
        "Integrated Science": 66.0,
        "Agriculture": 71.0,
        "Pre-Technical Studies": 50.0
      }
    },
    {
      "name": "ETHAN JOEL KAMAU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 24.0,
          "p2": 36.0,
          "pct": 60.0
        },
        "English": {
          "p1": 36.0,
          "p2": 30.0,
          "pct": 66.0
        },
        "Mathematics": 40.0,
        "Social Studies": 67.0,
        "Christian Religious Education": 66.0,
        "Creative Arts": 67.0,
        "Integrated Science": 71.0,
        "Agriculture": 71.0,
        "Pre-Technical Studies": 83.0
      }
    },
    {
      "name": "GABRIELLA MUMBI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 25.0,
          "p2": 33.0,
          "pct": 58.0
        },
        "English": {
          "p1": 43.0,
          "p2": 36.0,
          "pct": 79.0
        },
        "Mathematics": 61.0,
        "Social Studies": 88.0,
        "Christian Religious Education": 80.0,
        "Creative Arts": 77.0,
        "Integrated Science": 76.0,
        "Agriculture": 60.0,
        "Pre-Technical Studies": 83.0
      }
    },
    {
      "name": "JARED NJENGA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 19.0,
          "p2": 34.0,
          "pct": 53.0
        },
        "English": {
          "p1": 38.0,
          "p2": 31.0,
          "pct": 69.0
        },
        "Mathematics": 65.0,
        "Social Studies": 73.0,
        "Christian Religious Education": 62.0,
        "Creative Arts": 75.0,
        "Integrated Science": 84.0,
        "Agriculture": 73.0,
        "Pre-Technical Studies": 90.0
      }
    },
    {
      "name": "JAY NGUGI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 5.0,
          "p2": 21.0,
          "pct": 26.0
        },
        "English": {
          "p1": 29.0,
          "p2": 14.0,
          "pct": 43.0
        },
        "Mathematics": 52.0,
        "Social Studies": 45.0,
        "Christian Religious Education": 39.0,
        "Creative Arts": 54.0,
        "Integrated Science": 47.0,
        "Agriculture": 58.0,
        "Pre-Technical Studies": 61.0
      }
    },
    {
      "name": "JOY PRECIOUS WANJIRU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 47.0,
          "p2": 47.0,
          "pct": 94.0
        },
        "English": {
          "p1": 34.0,
          "p2": 44.0,
          "pct": 78.0
        },
        "Mathematics": 86.0,
        "Social Studies": 96.0,
        "Christian Religious Education": 88.0,
        "Creative Arts": 86.0,
        "Integrated Science": 84.0,
        "Agriculture": 83.0,
        "Pre-Technical Studies": 92.0
      }
    },
    {
      "name": "JOY SHARLEEN WAMBUI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 40.0,
          "p2": 44.0,
          "pct": 84.0
        },
        "English": {
          "p1": 45.0,
          "p2": 43.0,
          "pct": 88.0
        },
        "Mathematics": 88.0,
        "Social Studies": 98.0,
        "Christian Religious Education": 95.0,
        "Creative Arts": 90.0,
        "Integrated Science": 89.0,
        "Agriculture": 90.0,
        "Pre-Technical Studies": 91.0
      }
    },
    {
      "name": "LEON DAMIAN",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 19.0,
          "p2": 36.0,
          "pct": 55.0
        },
        "English": {
          "p1": 43.0,
          "p2": 39.0,
          "pct": 82.0
        },
        "Mathematics": 86.0,
        "Social Studies": 96.0,
        "Christian Religious Education": 68.0,
        "Creative Arts": 80.0,
        "Integrated Science": 79.0,
        "Agriculture": 80.0,
        "Pre-Technical Studies": 88.0
      }
    },
    {
      "name": "LOVELINE ANUPI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 27.0,
          "p2": 40.0,
          "pct": 67.0
        },
        "English": {
          "p1": 30.0,
          "p2": 28.0,
          "pct": 58.0
        },
        "Mathematics": 59.0,
        "Social Studies": 72.0,
        "Christian Religious Education": 49.0,
        "Creative Arts": 70.0,
        "Integrated Science": 56.0,
        "Agriculture": 74.0,
        "Pre-Technical Studies": 63.0
      }
    },
    {
      "name": "MEGAN NJERI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 39.0,
          "p2": 46.0,
          "pct": 85.0
        },
        "English": {
          "p1": 46.0,
          "p2": 39.0,
          "pct": 85.0
        },
        "Mathematics": 82.0,
        "Social Studies": 98.0,
        "Christian Religious Education": 94.0,
        "Creative Arts": 94.0,
        "Integrated Science": 87.0,
        "Agriculture": 80.0,
        "Pre-Technical Studies": 92.0
      }
    },
    {
      "name": "PATRIC MUKISA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 39.0,
          "p2": 48.0,
          "pct": 87.0
        },
        "English": {
          "p1": 41.0,
          "p2": 34.0,
          "pct": 75.0
        },
        "Mathematics": 88.0,
        "Social Studies": 89.0,
        "Christian Religious Education": 95.0,
        "Creative Arts": 85.0,
        "Integrated Science": 91.0,
        "Agriculture": 84.0,
        "Pre-Technical Studies": 83.0
      }
    },
    {
      "name": "REUBEN CHEGE",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 30.0,
          "p2": 41.0,
          "pct": 71.0
        },
        "English": {
          "p1": 26.0,
          "p2": 30.0,
          "pct": 56.0
        },
        "Mathematics": 59.0,
        "Social Studies": 80.0,
        "Christian Religious Education": 77.0,
        "Creative Arts": 64.0,
        "Integrated Science": 61.0,
        "Agriculture": 59.0,
        "Pre-Technical Studies": 71.0
      }
    },
    {
      "name": "RIC BRIAN IRUNGU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 11.0,
          "p2": 18.0,
          "pct": 29.0
        },
        "English": {
          "p1": 33.0,
          "p2": 31.0,
          "pct": 64.0
        },
        "Mathematics": 56.0,
        "Social Studies": 65.0,
        "Christian Religious Education": 45.0,
        "Creative Arts": 59.0,
        "Integrated Science": 59.0,
        "Agriculture": 56.0,
        "Pre-Technical Studies": 68.0
      }
    },
    {
      "name": "RYLAN GITHINJI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 7.0,
          "p2": 25.0,
          "pct": 32.0
        },
        "English": {
          "p1": 33.0,
          "p2": 28.0,
          "pct": 61.0
        },
        "Mathematics": 27.0,
        "Social Studies": 43.0,
        "Christian Religious Education": 61.0,
        "Creative Arts": 54.0,
        "Integrated Science": 60.0,
        "Agriculture": 51.0,
        "Pre-Technical Studies": 56.0
      }
    },
    {
      "name": "SHANTEL KURIA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 23.0,
          "p2": 34.0,
          "pct": 57.0
        },
        "English": {
          "p1": 42.0,
          "p2": 37.0,
          "pct": 79.0
        },
        "Mathematics": 58.0,
        "Social Studies": 86.0,
        "Christian Religious Education": 69.0,
        "Creative Arts": 79.0,
        "Integrated Science": 76.0,
        "Agriculture": 76.0,
        "Pre-Technical Studies": 83.0
      }
    },
    {
      "name": "TREVOR AMBUKA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 37.0,
          "p2": 44.0,
          "pct": 81.0
        },
        "English": {
          "p1": 41.0,
          "p2": 34.0,
          "pct": 75.0
        },
        "Mathematics": 74.0,
        "Social Studies": 85.0,
        "Christian Religious Education": 81.0,
        "Creative Arts": 69.0,
        "Integrated Science": 84.0,
        "Agriculture": 61.0,
        "Pre-Technical Studies": 84.0
      }
    },
    {
      "name": "SUBJECT RANK",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 8.0
        },
        "English": {
          "p1": null,
          "p2": null,
          "pct": 7.0
        },
        "Mathematics": 9.0,
        "Social Studies": 1.0,
        "Christian Religious Education": 6.0,
        "Creative Arts": 3.0,
        "Integrated Science": 4.0,
        "Agriculture": 5.0,
        "Pre-Technical Studies": 2.0
      }
    }
  ],
  "Grade 7": [
    {
      "name": "ABI SIFA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 38.0,
          "p2": 42.0,
          "pct": 80.0
        },
        "English": {
          "p1": 40.0,
          "p2": 39.0,
          "pct": 79.0
        },
        "Mathematics": 62.0,
        "Social Studies": 86.0,
        "Christian Religious Education": 88.0,
        "Creative Arts": 74.0,
        "Integrated Science": 86.0,
        "Agriculture": 78.0,
        "Pre-Technical Studies": 80.0
      }
    },
    {
      "name": "ABIGAEL BLESSING",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 25.0,
          "p2": 34.0,
          "pct": 59.0
        },
        "English": {
          "p1": 37.0,
          "p2": 33.0,
          "pct": 70.0
        },
        "Mathematics": 50.0,
        "Social Studies": 86.0,
        "Christian Religious Education": 86.0,
        "Creative Arts": 66.0,
        "Integrated Science": 74.0,
        "Agriculture": 84.0,
        "Pre-Technical Studies": 96.0
      }
    },
    {
      "name": "ABIGAIL WACUKA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 40.0,
          "p2": 40.0,
          "pct": 80.0
        },
        "English": {
          "p1": 45.0,
          "p2": 46.0,
          "pct": 91.0
        },
        "Mathematics": 82.0,
        "Social Studies": 96.0,
        "Christian Religious Education": 96.0,
        "Creative Arts": 86.0,
        "Integrated Science": 96.0,
        "Agriculture": 88.0,
        "Pre-Technical Studies": 90.0
      }
    },
    {
      "name": "ADRIAN MACHARIA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 8.0,
          "p2": 31.0,
          "pct": 39.0
        },
        "English": {
          "p1": 35.0,
          "p2": 36.0,
          "pct": 71.0
        },
        "Mathematics": 68.0,
        "Social Studies": 86.0,
        "Christian Religious Education": 64.0,
        "Creative Arts": 50.0,
        "Integrated Science": 86.0,
        "Agriculture": 68.0,
        "Pre-Technical Studies": 82.0
      }
    },
    {
      "name": "BECKY ACHANDO",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 27.0,
          "p2": 40.0,
          "pct": 67.0
        },
        "English": {
          "p1": 35.0,
          "p2": 35.0,
          "pct": 70.0
        },
        "Mathematics": 62.0,
        "Social Studies": 78.0,
        "Christian Religious Education": 66.0,
        "Creative Arts": 82.0,
        "Integrated Science": 74.0,
        "Agriculture": 56.0,
        "Pre-Technical Studies": 86.0
      }
    },
    {
      "name": "BRYAN NG'ANG'A",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 10.0,
          "p2": 33.0,
          "pct": 43.0
        },
        "English": {
          "p1": 40.0,
          "p2": 28.0,
          "pct": 68.0
        },
        "Mathematics": 44.0,
        "Social Studies": 46.0,
        "Christian Religious Education": 56.0,
        "Creative Arts": 60.0,
        "Integrated Science": 56.0,
        "Agriculture": 62.0,
        "Pre-Technical Studies": 76.0
      }
    },
    {
      "name": "ELI KARANJA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 16.0,
          "p2": 28.0,
          "pct": 44.0
        },
        "English": {
          "p1": 39.0,
          "p2": 42.0,
          "pct": 81.0
        },
        "Mathematics": 74.0,
        "Social Studies": 90.0,
        "Christian Religious Education": 76.0,
        "Creative Arts": 80.0,
        "Integrated Science": 90.0,
        "Agriculture": 82.0,
        "Pre-Technical Studies": 78.0
      }
    },
    {
      "name": "ELISHA BARAKA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 7.0,
          "p2": 14.0,
          "pct": 21.0
        },
        "English": {
          "p1": 32.0,
          "p2": 24.0,
          "pct": 56.0
        },
        "Mathematics": 50.0,
        "Social Studies": 70.0,
        "Christian Religious Education": 58.0,
        "Creative Arts": 60.0,
        "Integrated Science": 58.0,
        "Agriculture": 84.0,
        "Pre-Technical Studies": 66.0
      }
    },
    {
      "name": "EPHRAIM GITHINJI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 32.0,
          "p2": 35.0,
          "pct": 67.0
        },
        "English": {
          "p1": 38.0,
          "p2": 35.0,
          "pct": 73.0
        },
        "Mathematics": 38.0,
        "Social Studies": 74.0,
        "Christian Religious Education": 68.0,
        "Creative Arts": 62.0,
        "Integrated Science": 72.0,
        "Agriculture": 78.0,
        "Pre-Technical Studies": 76.0
      }
    },
    {
      "name": "ETHAN OTUNGA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 17.0,
          "p2": 38.0,
          "pct": 55.0
        },
        "English": {
          "p1": 38.0,
          "p2": 36.0,
          "pct": 74.0
        },
        "Mathematics": 70.0,
        "Social Studies": 80.0,
        "Christian Religious Education": 62.0,
        "Creative Arts": 58.0,
        "Integrated Science": 72.0,
        "Agriculture": 70.0,
        "Pre-Technical Studies": 78.0
      }
    },
    {
      "name": "JABARR MURIITHI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 17.0,
          "p2": 33.0,
          "pct": 50.0
        },
        "English": {
          "p1": 40.0,
          "p2": 37.0,
          "pct": 77.0
        },
        "Mathematics": 66.0,
        "Social Studies": 82.0,
        "Christian Religious Education": 74.0,
        "Creative Arts": 36.0,
        "Integrated Science": 84.0,
        "Agriculture": 78.0,
        "Pre-Technical Studies": 82.0
      }
    },
    {
      "name": "JAEL NG'ENDO",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 21.0,
          "p2": 33.0,
          "pct": 54.0
        },
        "English": {
          "p1": 35.0,
          "p2": 29.0,
          "pct": 64.0
        },
        "Mathematics": 46.0,
        "Social Studies": 72.0,
        "Christian Religious Education": 58.0,
        "Creative Arts": 60.0,
        "Integrated Science": 64.0,
        "Agriculture": 70.0,
        "Pre-Technical Studies": 76.0
      }
    },
    {
      "name": "JOEL MUTUA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 29.0,
          "p2": 37.0,
          "pct": 66.0
        },
        "English": {
          "p1": 38.0,
          "p2": 44.0,
          "pct": 82.0
        },
        "Mathematics": 86.0,
        "Social Studies": 94.0,
        "Christian Religious Education": 76.0,
        "Creative Arts": 82.0,
        "Integrated Science": 94.0,
        "Agriculture": 92.0,
        "Pre-Technical Studies": 84.0
      }
    },
    {
      "name": "JONATHAN MBURU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 17.0,
          "p2": 32.0,
          "pct": 49.0
        },
        "English": {
          "p1": 38.0,
          "p2": 36.0,
          "pct": 74.0
        },
        "Mathematics": 60.0,
        "Social Studies": 88.0,
        "Christian Religious Education": 82.0,
        "Creative Arts": 70.0,
        "Integrated Science": 84.0,
        "Agriculture": 82.0,
        "Pre-Technical Studies": 68.0
      }
    },
    {
      "name": "MERCY WANJIRU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 24.0,
          "p2": 37.0,
          "pct": 61.0
        },
        "English": {
          "p1": 35.0,
          "p2": 34.0,
          "pct": 69.0
        },
        "Mathematics": 78.0,
        "Social Studies": 147.0,
        "Christian Religious Education": 64.0,
        "Creative Arts": 74.0,
        "Integrated Science": 80.0,
        "Agriculture": 88.0,
        "Pre-Technical Studies": 78.0
      }
    },
    {
      "name": "NATASHA NJERI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 5.0,
          "p2": 21.0,
          "pct": 26.0
        },
        "English": {
          "p1": 28.0,
          "p2": 29.0,
          "pct": 57.0
        },
        "Mathematics": 32.0,
        "Social Studies": 42.0,
        "Christian Religious Education": 42.0,
        "Creative Arts": 36.0,
        "Integrated Science": 58.0,
        "Agriculture": 44.0,
        "Pre-Technical Studies": 58.0
      }
    },
    {
      "name": "NELSON KAMAU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 12.0,
          "p2": 31.0,
          "pct": 43.0
        },
        "English": {
          "p1": 35.0,
          "p2": 32.0,
          "pct": 67.0
        },
        "Mathematics": 52.0,
        "Social Studies": 66.0,
        "Christian Religious Education": 46.0,
        "Creative Arts": 54.0,
        "Integrated Science": 60.0,
        "Agriculture": 60.0,
        "Pre-Technical Studies": 80.0
      }
    },
    {
      "name": "NICHOLE NYOKABI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 13.0,
          "p2": 26.0,
          "pct": 39.0
        },
        "English": {
          "p1": 34.0,
          "p2": 33.0,
          "pct": 67.0
        },
        "Mathematics": 34.0,
        "Social Studies": 72.0,
        "Christian Religious Education": 68.0,
        "Creative Arts": 66.0,
        "Integrated Science": 60.0,
        "Agriculture": 54.0,
        "Pre-Technical Studies": 80.0
      }
    },
    {
      "name": "PETER NDUNG’U",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 26.0,
          "p2": 41.0,
          "pct": 67.0
        },
        "English": {
          "p1": 43.0,
          "p2": 44.0,
          "pct": 87.0
        },
        "Mathematics": 80.0,
        "Social Studies": 90.0,
        "Christian Religious Education": 64.0,
        "Creative Arts": 62.0,
        "Integrated Science": 92.0,
        "Agriculture": 88.0,
        "Pre-Technical Studies": 84.0
      }
    },
    {
      "name": "PURITY WANJIRU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 12.0,
          "p2": 31.0,
          "pct": 43.0
        },
        "English": {
          "p1": 35.0,
          "p2": 33.0,
          "pct": 68.0
        },
        "Mathematics": 24.0,
        "Social Studies": 86.0,
        "Christian Religious Education": 48.0,
        "Creative Arts": 52.0,
        "Integrated Science": 60.0,
        "Agriculture": 60.0,
        "Pre-Technical Studies": 76.0
      }
    },
    {
      "name": "RYAN GITAU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 37.0,
          "p2": 37.0,
          "pct": 74.0
        },
        "English": {
          "p1": 46.0,
          "p2": 43.0,
          "pct": 89.0
        },
        "Mathematics": 82.0,
        "Social Studies": 92.0,
        "Christian Religious Education": 88.0,
        "Creative Arts": 86.0,
        "Integrated Science": 94.0,
        "Agriculture": 84.0,
        "Pre-Technical Studies": 92.0
      }
    },
    {
      "name": "RYAN MUGAMBI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 40.0,
          "p2": 41.0,
          "pct": 81.0
        },
        "English": {
          "p1": 43.0,
          "p2": 42.0,
          "pct": 85.0
        },
        "Mathematics": 88.0,
        "Social Studies": 96.0,
        "Christian Religious Education": 88.0,
        "Creative Arts": 82.0,
        "Integrated Science": 90.0,
        "Agriculture": 88.0,
        "Pre-Technical Studies": 92.0
      }
    },
    {
      "name": "RYAN MUNYAO",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 31.0,
          "p2": 42.0,
          "pct": 73.0
        },
        "English": {
          "p1": 38.0,
          "p2": 38.0,
          "pct": 76.0
        },
        "Mathematics": 54.0,
        "Social Studies": 86.0,
        "Christian Religious Education": 90.0,
        "Creative Arts": 76.0,
        "Integrated Science": 78.0,
        "Agriculture": 80.0,
        "Pre-Technical Studies": 90.0
      }
    },
    {
      "name": "SHIREEN WANJIRU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 24.0,
          "p2": 38.0,
          "pct": 62.0
        },
        "English": {
          "p1": 40.0,
          "p2": 35.0,
          "pct": 75.0
        },
        "Mathematics": 72.0,
        "Social Studies": 78.0,
        "Christian Religious Education": 76.0,
        "Creative Arts": 62.0,
        "Integrated Science": 66.0,
        "Agriculture": 80.0,
        "Pre-Technical Studies": 92.0
      }
    },
    {
      "name": "TIFANNY NKATHA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 34.0,
          "p2": 37.0,
          "pct": 71.0
        },
        "English": {
          "p1": 40.0,
          "p2": 40.0,
          "pct": 80.0
        },
        "Mathematics": 74.0,
        "Social Studies": 84.0,
        "Christian Religious Education": 74.0,
        "Creative Arts": 72.0,
        "Integrated Science": 78.0,
        "Agriculture": 88.0,
        "Pre-Technical Studies": 90.0
      }
    },
    {
      "name": "TITO MOSES",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 11.0,
          "p2": 34.0,
          "pct": 45.0
        },
        "English": {
          "p1": 38.0,
          "p2": 32.0,
          "pct": 70.0
        },
        "Mathematics": 56.0,
        "Social Studies": 69.0,
        "Christian Religious Education": 62.0,
        "Creative Arts": 44.0,
        "Integrated Science": 78.0,
        "Agriculture": 74.0,
        "Pre-Technical Studies": 76.0
      }
    },
    {
      "name": "VELLA AKALA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 42.0,
          "p2": 49.0,
          "pct": 91.0
        },
        "English": {
          "p1": 39.0,
          "p2": 44.0,
          "pct": 83.0
        },
        "Mathematics": 68.0,
        "Social Studies": 92.0,
        "Christian Religious Education": 88.0,
        "Creative Arts": 78.0,
        "Integrated Science": 92.0,
        "Agriculture": 82.0,
        "Pre-Technical Studies": 92.0
      }
    },
    {
      "name": "SUBJECT RANK",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 9.0
        },
        "English": {
          "p1": null,
          "p2": null,
          "pct": 5.0
        },
        "Mathematics": 8.0,
        "Social Studies": 1.0,
        "Christian Religious Education": 6.0,
        "Creative Arts": 7.0,
        "Integrated Science": 3.0,
        "Agriculture": 4.0,
        "Pre-Technical Studies": 2.0
      }
    }
  ],
  "Grade 6": [
    {
      "name": "ABIGAIL WAITHIRA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 18.0,
          "p2": 4.0,
          "pct": 16.5
        },
        "English": {
          "p1": 25.0,
          "p2": 6.0,
          "pct": 23.25
        },
        "Mathematics": 24.0,
        "Social Studies": 16.0,
        "Christian Religious Education": 21.0,
        "Creative Arts": 24.0,
        "Science & Technology": 21.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "AIDEN GITONGAH",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 20.0,
          "p2": 3.0,
          "pct": 17.25
        },
        "English": {
          "p1": 22.0,
          "p2": 7.0,
          "pct": 21.75
        },
        "Mathematics": 17.0,
        "Social Studies": 21.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 25.0,
        "Science & Technology": 19.0,
        "Agriculture": 25.0
      }
    },
    {
      "name": "AITHAN NDARUA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 10.0,
          "p2": 3.0,
          "pct": 9.75
        },
        "English": {
          "p1": 25.0,
          "p2": 5.0,
          "pct": 22.5
        },
        "Mathematics": 14.0,
        "Social Studies": 12.0,
        "Christian Religious Education": 17.0,
        "Creative Arts": 20.0,
        "Science & Technology": 20.0,
        "Agriculture": 23.0
      }
    },
    {
      "name": "AMANDA WAITHERA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 17.0,
          "p2": 4.0,
          "pct": 15.75
        },
        "English": {
          "p1": 26.0,
          "p2": 6.0,
          "pct": 24.0
        },
        "Mathematics": 25.0,
        "Social Studies": 18.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 23.0,
        "Science & Technology": 22.0,
        "Agriculture": 24.0
      }
    },
    {
      "name": "ANAYA WANJIRU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 18.0,
          "p2": 6.0,
          "pct": 18.0
        },
        "English": {
          "p1": 25.0,
          "p2": 8.0,
          "pct": 24.75
        },
        "Mathematics": 25.0,
        "Social Studies": 20.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 25.0,
        "Science & Technology": 24.0,
        "Agriculture": 26.0
      }
    },
    {
      "name": "ARIANA BOSIBORI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 0.0,
          "p2": null,
          "pct": 0.0
        },
        "Mathematics": 0.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": null,
        "Science & Technology": null,
        "Agriculture": null
      }
    },
    {
      "name": "AYANNAH WANJRU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 12.0,
          "p2": 4.0,
          "pct": 12.0
        },
        "English": {
          "p1": 19.0,
          "p2": 6.0,
          "pct": 18.75
        },
        "Mathematics": 18.0,
        "Social Studies": 20.0,
        "Christian Religious Education": 23.0,
        "Creative Arts": 26.0,
        "Science & Technology": 19.0,
        "Agriculture": 21.0
      }
    },
    {
      "name": "BRANDON MWANIKI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 14.0,
          "p2": 5.0,
          "pct": 14.25
        },
        "English": {
          "p1": 23.0,
          "p2": 6.0,
          "pct": 21.75
        },
        "Mathematics": 22.0,
        "Social Studies": 21.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 26.0,
        "Science & Technology": 26.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "CRISTIN KIAMA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 15.0,
          "p2": 5.0,
          "pct": 15.0
        },
        "English": {
          "p1": 28.0,
          "p2": 6.0,
          "pct": 25.5
        },
        "Mathematics": 15.0,
        "Social Studies": 22.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 26.0,
        "Science & Technology": 24.0,
        "Agriculture": 23.0
      }
    },
    {
      "name": "DENICE CRYSTAL",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 16.0,
          "p2": 7.0,
          "pct": 17.25
        },
        "English": {
          "p1": 29.0,
          "p2": 8.0,
          "pct": 27.75
        },
        "Mathematics": 29.0,
        "Social Studies": 25.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 24.0,
        "Science & Technology": 25.0,
        "Agriculture": 26.0
      }
    },
    {
      "name": "ELLA KARIMI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 15.0,
          "p2": 3.0,
          "pct": 13.5
        },
        "English": {
          "p1": 22.0,
          "p2": 4.0,
          "pct": 19.5
        },
        "Mathematics": 14.0,
        "Social Studies": 12.0,
        "Christian Religious Education": 24.0,
        "Creative Arts": 20.0,
        "Science & Technology": 19.0,
        "Agriculture": 20.0
      }
    },
    {
      "name": "ETHAN NJUGUNA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 17.0,
          "p2": 5.0,
          "pct": 16.5
        },
        "English": {
          "p1": 28.0,
          "p2": 6.0,
          "pct": 25.5
        },
        "Mathematics": 21.0,
        "Social Studies": 25.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 25.0,
        "Science & Technology": 26.0,
        "Agriculture": 24.0
      }
    },
    {
      "name": "GLADWELL GAKII",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 19.0,
          "p2": 6.0,
          "pct": 18.75
        },
        "English": {
          "p1": 24.0,
          "p2": 5.0,
          "pct": 21.75
        },
        "Mathematics": 18.0,
        "Social Studies": 17.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 26.0,
        "Science & Technology": 25.0,
        "Agriculture": 29.0
      }
    },
    {
      "name": "IMANI MUNGUTI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 21.0,
          "p2": 5.0,
          "pct": 19.5
        },
        "English": {
          "p1": 23.0,
          "p2": 7.0,
          "pct": 22.5
        },
        "Mathematics": 18.0,
        "Social Studies": 20.0,
        "Christian Religious Education": 21.0,
        "Creative Arts": 23.0,
        "Science & Technology": 16.0,
        "Agriculture": 25.0
      }
    },
    {
      "name": "JAMHURI KUDA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 15.0,
          "p2": 4.0,
          "pct": 14.25
        },
        "English": {
          "p1": 26.0,
          "p2": 5.0,
          "pct": 23.25
        },
        "Mathematics": 19.0,
        "Social Studies": 22.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 26.0,
        "Science & Technology": 22.0,
        "Agriculture": 25.0
      }
    },
    {
      "name": "JAYDEN MANINI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 20.0,
          "p2": 3.0,
          "pct": 17.25
        },
        "English": {
          "p1": 25.0,
          "p2": 5.0,
          "pct": 22.5
        },
        "Mathematics": 22.0,
        "Social Studies": 18.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 25.0,
        "Science & Technology": 23.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "JENELLE MUTHONI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 18.0,
          "p2": 3.0,
          "pct": 15.75
        },
        "English": {
          "p1": 21.0,
          "p2": 8.0,
          "pct": 21.75
        },
        "Mathematics": 17.0,
        "Social Studies": 14.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 20.0,
        "Science & Technology": 21.0,
        "Agriculture": 20.0
      }
    },
    {
      "name": "LEO BROWN",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 17.0,
          "p2": 3.0,
          "pct": 15.0
        },
        "English": {
          "p1": 24.0,
          "p2": 6.0,
          "pct": 22.5
        },
        "Mathematics": 24.0,
        "Social Studies": 12.0,
        "Christian Religious Education": 21.0,
        "Creative Arts": 24.0,
        "Science & Technology": 19.0,
        "Agriculture": 19.0
      }
    },
    {
      "name": "LULU WAMBUI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 19.0,
          "p2": 6.0,
          "pct": 18.75
        },
        "English": {
          "p1": 25.0,
          "p2": 6.0,
          "pct": 23.25
        },
        "Mathematics": 25.0,
        "Social Studies": 20.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 27.0,
        "Science & Technology": 23.0,
        "Agriculture": 24.0
      }
    },
    {
      "name": "MELISSA MUGURE",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 16.0,
          "p2": 5.0,
          "pct": 15.75
        },
        "English": {
          "p1": 24.0,
          "p2": 6.0,
          "pct": 22.5
        },
        "Mathematics": 23.0,
        "Social Studies": 15.0,
        "Christian Religious Education": 24.0,
        "Creative Arts": 22.0,
        "Science & Technology": 23.0,
        "Agriculture": 26.0
      }
    },
    {
      "name": "MICHAEL MUKONO",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 8.0,
          "p2": 3.0,
          "pct": 8.25
        },
        "English": {
          "p1": 23.0,
          "p2": 5.0,
          "pct": 21.0
        },
        "Mathematics": 14.0,
        "Social Studies": 17.0,
        "Christian Religious Education": 23.0,
        "Creative Arts": 18.0,
        "Science & Technology": 19.0,
        "Agriculture": 16.0
      }
    },
    {
      "name": "MISHA SARAH",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 13.0,
          "p2": 4.0,
          "pct": 12.75
        },
        "English": {
          "p1": 22.0,
          "p2": 4.0,
          "pct": 19.5
        },
        "Mathematics": 12.0,
        "Social Studies": 13.0,
        "Christian Religious Education": 20.0,
        "Creative Arts": 24.0,
        "Science & Technology": 19.0,
        "Agriculture": 19.0
      }
    },
    {
      "name": "SAMUEL KIGATHE",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 11.0,
          "p2": 3.0,
          "pct": 10.5
        },
        "English": {
          "p1": 23.0,
          "p2": 5.0,
          "pct": 21.0
        },
        "Mathematics": 26.0,
        "Social Studies": 18.0,
        "Christian Religious Education": 21.0,
        "Creative Arts": 18.0,
        "Science & Technology": 22.0,
        "Agriculture": 16.0
      }
    },
    {
      "name": "SHANICE MURUGI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 14.0,
          "p2": 4.0,
          "pct": 13.5
        },
        "English": {
          "p1": 21.0,
          "p2": 7.0,
          "pct": 21.0
        },
        "Mathematics": 16.0,
        "Social Studies": 16.0,
        "Christian Religious Education": 21.0,
        "Creative Arts": 21.0,
        "Science & Technology": 23.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "TIMOTHY MBUGUA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 15.0,
          "p2": 3.0,
          "pct": 13.5
        },
        "English": {
          "p1": 26.0,
          "p2": 6.0,
          "pct": 24.0
        },
        "Mathematics": 21.0,
        "Social Studies": 18.0,
        "Christian Religious Education": 21.0,
        "Creative Arts": 26.0,
        "Science & Technology": 18.0,
        "Agriculture": 26.0
      }
    },
    {
      "name": "ZAWADI MAINA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 14.0,
          "p2": 4.0,
          "pct": 13.5
        },
        "English": {
          "p1": 27.0,
          "p2": 5.0,
          "pct": 24.0
        },
        "Mathematics": 19.0,
        "Social Studies": 16.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 24.0,
        "Science & Technology": 16.0,
        "Agriculture": 26.0
      }
    },
    {
      "name": "ZURI MUDAIDA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 19.0,
          "p2": 6.0,
          "pct": 18.75
        },
        "English": {
          "p1": 26.0,
          "p2": 6.0,
          "pct": 24.0
        },
        "Mathematics": 23.0,
        "Social Studies": 22.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 25.0,
        "Science & Technology": 24.0,
        "Agriculture": 23.0
      }
    },
    {
      "name": "SUBJECT RANK",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 8.0
        },
        "English": {
          "p1": null,
          "p2": null,
          "pct": 4.0
        },
        "Mathematics": 6.0,
        "Social Studies": 7.0,
        "Christian Religious Education": 1.0,
        "Creative Arts": 2.0,
        "Science & Technology": 5.0,
        "Agriculture": 3.0
      }
    }
  ],
  "Grade 5": [
    {
      "name": "ALYSSA WANGU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 21.0,
          "p2": 7.0,
          "pct": 21.0
        },
        "English": {
          "p1": 19.0,
          "p2": 6.0,
          "pct": 18.75
        },
        "Mathematics": 16.0,
        "Social Studies": 21.0,
        "Christian Religious Education": 12.0,
        "Creative Arts": 18.0,
        "Science & Technology": 17.0,
        "Agriculture": 18.0
      }
    },
    {
      "name": "COMARK ONANI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 14.0,
          "p2": 5.0,
          "pct": 14.25
        },
        "English": {
          "p1": 27.0,
          "p2": 6.0,
          "pct": 24.75
        },
        "Mathematics": 22.0,
        "Social Studies": 22.0,
        "Christian Religious Education": 21.0,
        "Creative Arts": 22.0,
        "Science & Technology": 26.0,
        "Agriculture": 26.0
      }
    },
    {
      "name": "DECLAN NGATIA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 15.0,
          "p2": 4.0,
          "pct": 14.25
        },
        "English": {
          "p1": 21.0,
          "p2": 5.0,
          "pct": 19.5
        },
        "Mathematics": 15.0,
        "Social Studies": 25.0,
        "Christian Religious Education": 14.0,
        "Creative Arts": 19.0,
        "Science & Technology": 18.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "DELVIN DUNCAN",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 19.0,
          "p2": 4.0,
          "pct": 17.25
        },
        "English": {
          "p1": 19.0,
          "p2": 6.0,
          "pct": 18.75
        },
        "Mathematics": 15.0,
        "Social Studies": 24.0,
        "Christian Religious Education": 15.0,
        "Creative Arts": 12.0,
        "Science & Technology": 14.0,
        "Agriculture": 23.0
      }
    },
    {
      "name": "ELAM CADE",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 19.0,
          "p2": 5.0,
          "pct": 18.0
        },
        "English": {
          "p1": 23.0,
          "p2": 6.0,
          "pct": 21.75
        },
        "Mathematics": 20.0,
        "Social Studies": 21.0,
        "Christian Religious Education": 21.0,
        "Creative Arts": 25.0,
        "Science & Technology": 28.0,
        "Agriculture": 25.0
      }
    },
    {
      "name": "ELLA IMANI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 22.0,
          "p2": 6.0,
          "pct": 21.0
        },
        "English": {
          "p1": 21.0,
          "p2": 7.0,
          "pct": 21.0
        },
        "Mathematics": 12.0,
        "Social Studies": 25.0,
        "Christian Religious Education": 20.0,
        "Creative Arts": 26.0,
        "Science & Technology": 17.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "ELSIE AGNES",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 20.0,
          "p2": 7.0,
          "pct": 20.25
        },
        "English": {
          "p1": 24.0,
          "p2": 8.0,
          "pct": 24.0
        },
        "Mathematics": 19.0,
        "Social Studies": 30.0,
        "Christian Religious Education": 21.0,
        "Creative Arts": 27.0,
        "Science & Technology": 29.0,
        "Agriculture": 29.0
      }
    },
    {
      "name": "ETHAN MURUTHI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 12.0,
          "p2": 4.0,
          "pct": 12.0
        },
        "English": {
          "p1": 18.0,
          "p2": 6.0,
          "pct": 18.0
        },
        "Mathematics": 10.0,
        "Social Studies": 26.0,
        "Christian Religious Education": 23.0,
        "Creative Arts": 13.0,
        "Science & Technology": 23.0,
        "Agriculture": 24.0
      }
    },
    {
      "name": "ETHAN SAITEMU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 7.0,
          "p2": 4.0,
          "pct": 8.25
        },
        "English": {
          "p1": 15.0,
          "p2": 4.0,
          "pct": 14.25
        },
        "Mathematics": 12.0,
        "Social Studies": 20.0,
        "Christian Religious Education": 9.0,
        "Creative Arts": 18.0,
        "Science & Technology": 19.0,
        "Agriculture": 17.0
      }
    },
    {
      "name": "JANE WANJIRU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 14.0,
          "p2": 3.0,
          "pct": 12.75
        },
        "English": {
          "p1": 19.0,
          "p2": 6.0,
          "pct": 18.75
        },
        "Mathematics": 12.0,
        "Social Studies": 13.0,
        "Christian Religious Education": 14.0,
        "Creative Arts": 15.0,
        "Science & Technology": 17.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "JASON THUO",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 11.0,
          "p2": 5.0,
          "pct": 12.0
        },
        "English": {
          "p1": 22.0,
          "p2": 6.0,
          "pct": 21.0
        },
        "Mathematics": 18.0,
        "Social Studies": 23.0,
        "Christian Religious Education": 24.0,
        "Creative Arts": 22.0,
        "Science & Technology": 26.0,
        "Agriculture": 28.0
      }
    },
    {
      "name": "JEREMIAH KEITH",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 12.0,
          "p2": 3.0,
          "pct": 11.25
        },
        "English": {
          "p1": 17.0,
          "p2": 4.0,
          "pct": 15.75
        },
        "Mathematics": 15.0,
        "Social Studies": 23.0,
        "Christian Religious Education": 12.0,
        "Creative Arts": 16.0,
        "Science & Technology": 18.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "JEROME MWANGI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 20.0,
          "p2": 6.0,
          "pct": 19.5
        },
        "English": {
          "p1": 21.0,
          "p2": 6.0,
          "pct": 20.25
        },
        "Mathematics": 20.0,
        "Social Studies": 29.0,
        "Christian Religious Education": 20.0,
        "Creative Arts": 27.0,
        "Science & Technology": 26.0,
        "Agriculture": 28.0
      }
    },
    {
      "name": "JOSEPH KARANJA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 18.0,
          "p2": 5.0,
          "pct": 17.25
        },
        "English": {
          "p1": 20.0,
          "p2": 6.0,
          "pct": 19.5
        },
        "Mathematics": 26.0,
        "Social Studies": 24.0,
        "Christian Religious Education": 12.0,
        "Creative Arts": 24.0,
        "Science & Technology": 19.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "LIAM TYLER",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 15.0,
          "p2": 6.0,
          "pct": 15.75
        },
        "English": {
          "p1": 21.0,
          "p2": 5.0,
          "pct": 19.5
        },
        "Mathematics": 24.0,
        "Social Studies": 27.0,
        "Christian Religious Education": 21.0,
        "Creative Arts": 20.0,
        "Science & Technology": 20.0,
        "Agriculture": 25.0
      }
    },
    {
      "name": "MELVIN KARANI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 15.0,
          "p2": 6.0,
          "pct": 15.75
        },
        "English": {
          "p1": 26.0,
          "p2": 8.0,
          "pct": 25.5
        },
        "Mathematics": 21.0,
        "Social Studies": 22.0,
        "Christian Religious Education": 20.0,
        "Creative Arts": 25.0,
        "Science & Technology": 17.0,
        "Agriculture": 25.0
      }
    },
    {
      "name": "METOYA WARUIRU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 8.0,
          "p2": 4.0,
          "pct": 9.0
        },
        "English": {
          "p1": 25.0,
          "p2": 5.0,
          "pct": 22.5
        },
        "Mathematics": 11.0,
        "Social Studies": 23.0,
        "Christian Religious Education": 18.0,
        "Creative Arts": 21.0,
        "Science & Technology": 17.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "OCHENGO OCHENGO",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 14.0,
          "p2": 5.0,
          "pct": 14.25
        },
        "English": {
          "p1": 20.0,
          "p2": 6.0,
          "pct": 19.5
        },
        "Mathematics": 15.0,
        "Social Studies": 23.0,
        "Christian Religious Education": 15.0,
        "Creative Arts": 15.0,
        "Science & Technology": 20.0,
        "Agriculture": 19.0
      }
    },
    {
      "name": "PRINCE CHEGE",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 15.0,
          "p2": 5.0,
          "pct": 15.0
        },
        "English": {
          "p1": 23.0,
          "p2": 6.0,
          "pct": 21.75
        },
        "Mathematics": 16.0,
        "Social Studies": 24.0,
        "Christian Religious Education": 24.0,
        "Creative Arts": 24.0,
        "Science & Technology": 23.0,
        "Agriculture": 24.0
      }
    },
    {
      "name": "ROSE ISIS",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 10.0,
          "p2": 4.0,
          "pct": 10.5
        },
        "English": {
          "p1": 21.0,
          "p2": 6.0,
          "pct": 20.25
        },
        "Mathematics": 14.0,
        "Social Studies": 20.0,
        "Christian Religious Education": 12.0,
        "Creative Arts": 17.0,
        "Science & Technology": 17.0,
        "Agriculture": 23.0
      }
    },
    {
      "name": "SCOTT GATU",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 6.0,
          "p2": 3.0,
          "pct": 6.75
        },
        "English": {
          "p1": 18.0,
          "p2": 4.0,
          "pct": 16.5
        },
        "Mathematics": 12.0,
        "Social Studies": 21.0,
        "Christian Religious Education": 12.0,
        "Creative Arts": 16.0,
        "Science & Technology": 13.0,
        "Agriculture": 14.0
      }
    },
    {
      "name": "SIMON KAMAU GICHANE",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 20.0,
          "p2": 6.0,
          "pct": 19.5
        },
        "English": {
          "p1": 24.0,
          "p2": 6.0,
          "pct": 22.5
        },
        "Mathematics": 22.0,
        "Social Studies": 28.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 26.0,
        "Science & Technology": 24.0,
        "Agriculture": 24.0
      }
    },
    {
      "name": "TACARI GENESIS",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 14.0,
          "p2": 4.0,
          "pct": 13.5
        },
        "English": {
          "p1": 20.0,
          "p2": 5.0,
          "pct": 18.75
        },
        "Mathematics": 14.0,
        "Social Studies": 23.0,
        "Christian Religious Education": 17.0,
        "Creative Arts": 23.0,
        "Science & Technology": 19.0,
        "Agriculture": 26.0
      }
    },
    {
      "name": "TALEK THINWA",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 7.0,
          "p2": 3.0,
          "pct": 7.5
        },
        "English": {
          "p1": 22.0,
          "p2": 6.0,
          "pct": 21.0
        },
        "Mathematics": 15.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": null,
        "Science & Technology": 28.0,
        "Agriculture": 28.0
      }
    },
    {
      "name": "TANA KAGAI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 6.0,
          "p2": 3.0,
          "pct": 6.75
        },
        "English": {
          "p1": 20.0,
          "p2": 6.0,
          "pct": 19.5
        },
        "Mathematics": 14.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": null,
        "Science & Technology": 26.0,
        "Agriculture": 26.0
      }
    },
    {
      "name": "WALTER CHEGE",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 14.0,
          "p2": 4.0,
          "pct": 13.5
        },
        "English": {
          "p1": 21.0,
          "p2": 5.0,
          "pct": 19.5
        },
        "Mathematics": 13.0,
        "Social Studies": 22.0,
        "Christian Religious Education": 18.0,
        "Creative Arts": 19.0,
        "Science & Technology": 19.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "XARIA NJERI",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": 16.0,
          "p2": 5.0,
          "pct": 15.75
        },
        "English": {
          "p1": 24.0,
          "p2": 7.0,
          "pct": 23.25
        },
        "Mathematics": 19.0,
        "Social Studies": 25.0,
        "Christian Religious Education": 17.0,
        "Creative Arts": 20.0,
        "Science & Technology": 24.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "SUBJECT RANK",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 8.0
        },
        "English": {
          "p1": null,
          "p2": null,
          "pct": 4.0
        },
        "Mathematics": 6.0,
        "Social Studies": 2.0,
        "Christian Religious Education": 7.0,
        "Creative Arts": 5.0,
        "Science & Technology": 3.0,
        "Agriculture": 1.0
      }
    }
  ],
  "Grade 4": [
    {
      "name": "Adrian wanjohi",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 16.0,
          "p2": 5.0,
          "pct": 15.75
        },
        "English": {
          "p1": 14.0,
          "p2": 7.0,
          "pct": 15.75
        },
        "Mathematics": 11.0,
        "Social Studies": 26.0,
        "Christian Religious Education": 15.0,
        "Creative Arts": 22.0,
        "Science & Technology": 23.0,
        "Agriculture": 21.0
      }
    },
    {
      "name": "Alma wanjiru",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 17.0,
          "p2": 7.0,
          "pct": 18.0
        },
        "English": {
          "p1": 23.0,
          "p2": 7.0,
          "pct": 22.5
        },
        "Mathematics": 24.0,
        "Social Studies": 26.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 24.0,
        "Science & Technology": 17.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "Amanda Faith",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 21.0,
          "p2": 4.0,
          "pct": 18.75
        },
        "Mathematics": 24.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 24.0,
        "Science & Technology": 24.0,
        "Agriculture": null
      }
    },
    {
      "name": "Annika chibai",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 21.0,
          "p2": 6.0,
          "pct": 20.25
        },
        "English": {
          "p1": 21.0,
          "p2": 7.0,
          "pct": 21.0
        },
        "Mathematics": 23.0,
        "Social Studies": 26.0,
        "Christian Religious Education": 21.0,
        "Creative Arts": 28.0,
        "Science & Technology": 24.0,
        "Agriculture": 23.0
      }
    },
    {
      "name": "Azariah Ndwiga",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 12.0,
          "p2": 6.0,
          "pct": 13.5
        },
        "English": {
          "p1": 6.0,
          "p2": 5.0,
          "pct": 8.25
        },
        "Mathematics": 14.0,
        "Social Studies": 20.0,
        "Christian Religious Education": 12.0,
        "Creative Arts": 23.0,
        "Science & Technology": 25.0,
        "Agriculture": 21.0
      }
    },
    {
      "name": "Azariah wega",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 25.0,
          "p2": 8.0,
          "pct": 24.75
        },
        "English": {
          "p1": 28.0,
          "p2": 8.0,
          "pct": 27.0
        },
        "Mathematics": 30.0,
        "Social Studies": 29.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 28.0,
        "Science & Technology": 29.0,
        "Agriculture": 29.0
      }
    },
    {
      "name": "Baraka Mutuma",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 15.0,
          "p2": 5.0,
          "pct": 15.0
        },
        "English": {
          "p1": 22.0,
          "p2": 6.0,
          "pct": 21.0
        },
        "Mathematics": 24.0,
        "Social Studies": 26.0,
        "Christian Religious Education": 18.0,
        "Creative Arts": 23.0,
        "Science & Technology": 24.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "Brayden Murigi",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 18.0,
          "p2": 7.0,
          "pct": 18.75
        },
        "English": {
          "p1": 24.0,
          "p2": 6.0,
          "pct": 22.5
        },
        "Mathematics": 25.0,
        "Social Studies": 28.0,
        "Christian Religious Education": 23.0,
        "Creative Arts": 25.0,
        "Science & Technology": 26.0,
        "Agriculture": 25.0
      }
    },
    {
      "name": "Bryden Kabiru",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 15.0,
          "p2": 5.0,
          "pct": 15.0
        },
        "English": {
          "p1": 16.0,
          "p2": 6.0,
          "pct": 16.5
        },
        "Mathematics": 21.0,
        "Social Studies": 23.0,
        "Christian Religious Education": 12.0,
        "Creative Arts": 19.0,
        "Science & Technology": 23.0,
        "Agriculture": 25.0
      }
    },
    {
      "name": "Bryson Nganga",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 22.0,
          "p2": 5.0,
          "pct": 20.25
        },
        "Mathematics": 18.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 24.0,
        "Science & Technology": 25.0,
        "Agriculture": 23.0
      }
    },
    {
      "name": "Christian Ndungu",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 14.0,
          "p2": 5.0,
          "pct": 14.25
        },
        "Mathematics": 19.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 24.0,
        "Science & Technology": 19.0,
        "Agriculture": 21.0
      }
    },
    {
      "name": "Christian Taji",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 21.0,
          "p2": 6.0,
          "pct": 20.25
        },
        "Mathematics": 26.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 26.0,
        "Science & Technology": 25.0,
        "Agriculture": 24.0
      }
    },
    {
      "name": "Crevis Ramah Ngobe",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 12.0,
          "p2": 3.0,
          "pct": 11.25
        },
        "English": {
          "p1": 6.0,
          "p2": 6.0,
          "pct": 9.0
        },
        "Mathematics": 12.0,
        "Social Studies": 18.0,
        "Christian Religious Education": 14.0,
        "Creative Arts": 18.0,
        "Science & Technology": 13.0,
        "Agriculture": 17.0
      }
    },
    {
      "name": "Danniyal Hassan",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 12.0,
          "p2": 6.0,
          "pct": 13.5
        },
        "English": {
          "p1": 14.0,
          "p2": 7.0,
          "pct": 15.75
        },
        "Mathematics": 17.0,
        "Social Studies": 25.0,
        "Christian Religious Education": 9.0,
        "Creative Arts": 22.0,
        "Science & Technology": 23.0,
        "Agriculture": 18.0
      }
    },
    {
      "name": "Darel Kiki",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 18.0,
          "p2": 6.0,
          "pct": 18.0
        },
        "Mathematics": 17.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 23.0,
        "Science & Technology": 24.0,
        "Agriculture": 25.0
      }
    },
    {
      "name": "Dylan Gichuki",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 14.0,
          "p2": 5.0,
          "pct": 14.25
        },
        "English": {
          "p1": 12.0,
          "p2": 6.0,
          "pct": 13.5
        },
        "Mathematics": 15.0,
        "Social Studies": 21.0,
        "Christian Religious Education": 12.0,
        "Creative Arts": 22.0,
        "Science & Technology": 18.0,
        "Agriculture": 17.0
      }
    },
    {
      "name": "Dylan Maina",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 6.0,
          "p2": 2.0,
          "pct": 6.0
        },
        "Mathematics": 13.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 7.0,
        "Science & Technology": 14.0,
        "Agriculture": 13.0
      }
    },
    {
      "name": "Ellah Muigai",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 20.0,
          "p2": 6.0,
          "pct": 19.5
        },
        "English": {
          "p1": 17.0,
          "p2": 7.0,
          "pct": 18.0
        },
        "Mathematics": 14.0,
        "Social Studies": 22.0,
        "Christian Religious Education": 12.0,
        "Creative Arts": 27.0,
        "Science & Technology": 18.0,
        "Agriculture": 23.0
      }
    },
    {
      "name": "Elysia njoki",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 20.0,
          "p2": 5.0,
          "pct": 18.75
        },
        "English": {
          "p1": 17.0,
          "p2": 7.0,
          "pct": 18.0
        },
        "Mathematics": 16.0,
        "Social Studies": 25.0,
        "Christian Religious Education": 9.0,
        "Creative Arts": 23.0,
        "Science & Technology": 17.0,
        "Agriculture": 14.0
      }
    },
    {
      "name": "Emmanuella Njoki",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 24.0,
          "p2": 8.0,
          "pct": 24.0
        },
        "Mathematics": 25.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 25.0,
        "Science & Technology": 28.0,
        "Agriculture": 25.0
      }
    },
    {
      "name": "Ethan Mwangi",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 21.0,
          "p2": 6.0,
          "pct": 20.25
        },
        "English": {
          "p1": 19.0,
          "p2": 7.0,
          "pct": 19.5
        },
        "Mathematics": 19.0,
        "Social Studies": 25.0,
        "Christian Religious Education": 15.0,
        "Creative Arts": 23.0,
        "Science & Technology": 28.0,
        "Agriculture": 23.0
      }
    },
    {
      "name": "Farell Terrence",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 17.0,
          "p2": 6.0,
          "pct": 17.25
        },
        "English": {
          "p1": 22.0,
          "p2": 8.0,
          "pct": 22.5
        },
        "Mathematics": 24.0,
        "Social Studies": 23.0,
        "Christian Religious Education": 20.0,
        "Creative Arts": 24.0,
        "Science & Technology": 25.0,
        "Agriculture": 25.0
      }
    },
    {
      "name": "Gift Kamau",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 10.0,
          "p2": 5.0,
          "pct": 11.25
        },
        "Mathematics": 12.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 17.0,
        "Science & Technology": 22.0,
        "Agriculture": 13.0
      }
    },
    {
      "name": "Jameliah Njura",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 14.0,
          "p2": 5.0,
          "pct": 14.25
        },
        "English": {
          "p1": 14.0,
          "p2": 7.0,
          "pct": 15.75
        },
        "Mathematics": 13.0,
        "Social Studies": 22.0,
        "Christian Religious Education": 14.0,
        "Creative Arts": 26.0,
        "Science & Technology": 18.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "Janice Wambui",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 28.0,
          "p2": 8.0,
          "pct": 27.0
        },
        "Mathematics": 27.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 28.0,
        "Science & Technology": 26.0,
        "Agriculture": 25.0
      }
    },
    {
      "name": "Jason Luke",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 12.0,
          "p2": 5.0,
          "pct": 12.75
        },
        "English": {
          "p1": 17.0,
          "p2": 7.0,
          "pct": 18.0
        },
        "Mathematics": 20.0,
        "Social Studies": 27.0,
        "Christian Religious Education": 15.0,
        "Creative Arts": 24.0,
        "Science & Technology": 24.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "Jayden Joash",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 11.0,
          "p2": 5.0,
          "pct": 12.0
        },
        "Mathematics": 8.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 15.0,
        "Science & Technology": 19.0,
        "Agriculture": 12.0
      }
    },
    {
      "name": "Jayden Mutimu",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 13.0,
          "p2": 2.0,
          "pct": 11.25
        },
        "Mathematics": 11.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 18.0,
        "Science & Technology": 16.0,
        "Agriculture": 20.0
      }
    },
    {
      "name": "John Ngugi",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 13.0,
          "p2": 6.0,
          "pct": 14.25
        },
        "English": {
          "p1": 22.0,
          "p2": 4.0,
          "pct": 19.5
        },
        "Mathematics": 21.0,
        "Social Studies": 23.0,
        "Christian Religious Education": 14.0,
        "Creative Arts": 26.0,
        "Science & Technology": 19.0,
        "Agriculture": 21.0
      }
    },
    {
      "name": "Kayla Nduta",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 14.0,
          "p2": 5.0,
          "pct": 14.25
        },
        "English": {
          "p1": 23.0,
          "p2": 8.0,
          "pct": 23.25
        },
        "Mathematics": 22.0,
        "Social Studies": 25.0,
        "Christian Religious Education": 18.0,
        "Creative Arts": 26.0,
        "Science & Technology": 20.0,
        "Agriculture": 20.0
      }
    },
    {
      "name": "Ken Samuel wagereka",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 16.0,
          "p2": 5.0,
          "pct": 15.75
        },
        "English": {
          "p1": 18.0,
          "p2": 6.0,
          "pct": 18.0
        },
        "Mathematics": 22.0,
        "Social Studies": 27.0,
        "Christian Religious Education": 17.0,
        "Creative Arts": 24.0,
        "Science & Technology": 23.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "Leticia Wanjiru",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 16.0,
          "p2": 5.0,
          "pct": 15.75
        },
        "Mathematics": 14.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 18.0,
        "Science & Technology": 18.0,
        "Agriculture": 24.0
      }
    },
    {
      "name": "Merry Natalie Tatiana",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 20.0,
          "p2": 7.0,
          "pct": 20.25
        },
        "Mathematics": 21.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 23.0,
        "Science & Technology": 26.0,
        "Agriculture": 23.0
      }
    },
    {
      "name": "Mikaella Rehema",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 26.0,
          "p2": 6.0,
          "pct": 24.0
        },
        "Mathematics": 20.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 25.0,
        "Science & Technology": 28.0,
        "Agriculture": 28.0
      }
    },
    {
      "name": "Morgan kamau",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 17.0,
          "p2": 6.0,
          "pct": 17.25
        },
        "English": {
          "p1": 21.0,
          "p2": 6.0,
          "pct": 20.25
        },
        "Mathematics": 22.0,
        "Social Studies": 26.0,
        "Christian Religious Education": 20.0,
        "Creative Arts": 26.0,
        "Science & Technology": 26.0,
        "Agriculture": 25.0
      }
    },
    {
      "name": "Nadia chebet",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 20.0,
          "p2": 6.0,
          "pct": 19.5
        },
        "English": {
          "p1": 19.0,
          "p2": 8.0,
          "pct": 20.25
        },
        "Mathematics": 21.0,
        "Social Studies": 22.0,
        "Christian Religious Education": 15.0,
        "Creative Arts": 27.0,
        "Science & Technology": 22.0,
        "Agriculture": 28.0
      }
    },
    {
      "name": "Natasha Wambui",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 10.0,
          "p2": 5.0,
          "pct": 11.25
        },
        "Mathematics": 17.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 19.0,
        "Science & Technology": 20.0,
        "Agriculture": 21.0
      }
    },
    {
      "name": "Nathan Njagi",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 13.0,
          "p2": 5.0,
          "pct": 13.5
        },
        "English": {
          "p1": 21.0,
          "p2": 4.0,
          "pct": 18.75
        },
        "Mathematics": 23.0,
        "Social Studies": 23.0,
        "Christian Religious Education": 20.0,
        "Creative Arts": 24.0,
        "Science & Technology": 24.0,
        "Agriculture": 17.0
      }
    },
    {
      "name": "Nicole Nyambura",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 23.0,
          "p2": 6.0,
          "pct": 21.75
        },
        "Mathematics": 23.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 18.0,
        "Science & Technology": 22.0,
        "Agriculture": 20.0
      }
    },
    {
      "name": "Renee Sharlotte",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 19.0,
          "p2": 7.0,
          "pct": 19.5
        },
        "Mathematics": 10.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 17.0,
        "Science & Technology": 22.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "Rhema Wandia",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 20.0,
          "p2": 5.0,
          "pct": 18.75
        },
        "Mathematics": 13.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 23.0,
        "Science & Technology": 20.0,
        "Agriculture": 24.0
      }
    },
    {
      "name": "Ryan Muuo",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 20.0,
          "p2": 6.0,
          "pct": 19.5
        },
        "Mathematics": 15.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 22.0,
        "Science & Technology": 25.0,
        "Agriculture": 25.0
      }
    },
    {
      "name": "Shem Muriaigiri",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 20.0,
          "p2": 5.0,
          "pct": 18.75
        },
        "Mathematics": 24.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 23.0,
        "Science & Technology": 17.0,
        "Agriculture": 23.0
      }
    },
    {
      "name": "Shimeah Mutugi",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 15.0,
          "p2": 5.0,
          "pct": 15.0
        },
        "English": {
          "p1": 22.0,
          "p2": 7.0,
          "pct": 21.75
        },
        "Mathematics": 19.0,
        "Social Studies": 26.0,
        "Christian Religious Education": 20.0,
        "Creative Arts": 24.0,
        "Science & Technology": 29.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "Sifa Wanjiru",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 17.0,
          "p2": 5.0,
          "pct": 16.5
        },
        "English": {
          "p1": 11.0,
          "p2": 6.0,
          "pct": 12.75
        },
        "Mathematics": 16.0,
        "Social Studies": 25.0,
        "Christian Religious Education": 17.0,
        "Creative Arts": 24.0,
        "Science & Technology": 19.0,
        "Agriculture": 18.0
      }
    },
    {
      "name": "Stefan Gathagu",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 26.0,
          "p2": 7.0,
          "pct": 24.75
        },
        "Mathematics": 26.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 25.0,
        "Science & Technology": 28.0,
        "Agriculture": 26.0
      }
    },
    {
      "name": "Steve Benson",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 19.0,
          "p2": 7.0,
          "pct": 19.5
        },
        "Mathematics": 23.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 24.0,
        "Science & Technology": 22.0,
        "Agriculture": 24.0
      }
    },
    {
      "name": "Teyo Matalanga",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 20.0,
          "p2": 7.0,
          "pct": 20.25
        },
        "Mathematics": 15.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 19.0,
        "Science & Technology": 23.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "Tyrell Kaumah",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 13.0,
          "p2": 7.0,
          "pct": 15.0
        },
        "English": {
          "p1": 18.0,
          "p2": 5.0,
          "pct": 17.25
        },
        "Mathematics": 21.0,
        "Social Studies": 27.0,
        "Christian Religious Education": 20.0,
        "Creative Arts": 26.0,
        "Science & Technology": 24.0,
        "Agriculture": 24.0
      }
    },
    {
      "name": "Victor Morgan",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 17.0,
          "p2": 5.0,
          "pct": 16.5
        },
        "Mathematics": 11.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 22.0,
        "Science & Technology": 19.0,
        "Agriculture": 24.0
      }
    },
    {
      "name": "Zani Nganga",
      "stream": "YELLOW",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 0.0
        },
        "English": {
          "p1": 21.0,
          "p2": 5.0,
          "pct": 19.5
        },
        "Mathematics": 25.0,
        "Social Studies": null,
        "Christian Religious Education": null,
        "Creative Arts": 22.0,
        "Science & Technology": 23.0,
        "Agriculture": 25.0
      }
    },
    {
      "name": "Zemirah Zuwema",
      "stream": "RED",
      "marks": {
        "Kiswahili": {
          "p1": 21.0,
          "p2": 8.0,
          "pct": 21.75
        },
        "English": {
          "p1": 22.0,
          "p2": 7.0,
          "pct": 21.75
        },
        "Mathematics": 24.0,
        "Social Studies": 25.0,
        "Christian Religious Education": 24.0,
        "Creative Arts": 28.0,
        "Science & Technology": 23.0,
        "Agriculture": 22.0
      }
    },
    {
      "name": "SUBJECT RANK",
      "stream": "",
      "marks": {
        "Kiswahili": {
          "p1": null,
          "p2": null,
          "pct": 8.0
        },
        "English": {
          "p1": null,
          "p2": null,
          "pct": 5.0
        },
        "Mathematics": 4.0,
        "Social Studies": 6.0,
        "Christian Religious Education": 7.0,
        "Creative Arts": 1.0,
        "Science & Technology": 2.0,
        "Agriculture": 3.0
      }
    }
  ],
  "Grade 3": [
    {
      "name": "ADRIEL SAMUEL",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 27.0,
        "English": 29.0,
        "Kiswahili": 29.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "ALBA SELAH",
      "stream": "RED",
      "marks": {
        "Mathematics": 22.0,
        "English": 30.0,
        "Kiswahili": 29.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "AVERY ASIYO",
      "stream": "RED",
      "marks": {
        "Mathematics": 29.0,
        "English": 28.0,
        "Kiswahili": 26.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "BERNICE WANGU AKUBE",
      "stream": "RED",
      "marks": {
        "Mathematics": 26.0,
        "English": 27.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "CALVIN MAINA",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": null,
        "English": null,
        "Kiswahili": null,
        "Environmental Activities": null,
        "Christian Religious Education": null,
        "Creative Arts": null
      }
    },
    {
      "name": "CELINE MUGURE",
      "stream": "RED",
      "marks": {
        "Mathematics": 23.0,
        "English": 28.0,
        "Kiswahili": 28.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 22.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "CHARLOTTE WANJA",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 27.0,
        "English": 29.0,
        "Kiswahili": 25.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "ELLAH NJERI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 17.0,
        "English": 26.0,
        "Kiswahili": 15.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 18.0,
        "Creative Arts": 24.0
      }
    },
    {
      "name": "FABIANA NAMUNYAK",
      "stream": "RED",
      "marks": {
        "Mathematics": 25.0,
        "English": 30.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "GABRIEL MACHARIA",
      "stream": "RED",
      "marks": {
        "Mathematics": 29.0,
        "English": 28.0,
        "Kiswahili": 26.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "GABRIELLA WAMBUI",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "GEORGE MBUGUA",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 30.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "IVY MUKUHI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 29.0,
        "English": 30.0,
        "Kiswahili": 29.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "JANAYA WAMBUI",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 30.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "JARED WAIGURU",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 26.0,
        "English": 29.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "JASON MUCHOKI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 28.0,
        "English": 30.0,
        "Kiswahili": 28.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "JASON PRINCE",
      "stream": "RED",
      "marks": {
        "Mathematics": 28.0,
        "English": 30.0,
        "Kiswahili": 26.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "JEREMY KIAMA",
      "stream": "RED",
      "marks": {
        "Mathematics": 29.0,
        "English": 29.0,
        "Kiswahili": 24.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "JOHN MUNGAI",
      "stream": "RED",
      "marks": {
        "Mathematics": 29.0,
        "English": 29.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "JONATHAN NDIRANGU",
      "stream": "RED",
      "marks": {
        "Mathematics": 12.0,
        "English": 21.0,
        "Kiswahili": 12.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 18.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "KIMATHI GATHURA",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 29.0,
        "Kiswahili": 26.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "LEILA KIBE",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 24.0,
        "English": 25.0,
        "Kiswahili": 17.0,
        "Environmental Activities": 25.0,
        "Christian Religious Education": 18.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "LEILANI SAYO",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 21.0,
        "English": 28.0,
        "Kiswahili": 24.0,
        "Environmental Activities": 25.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "LEVI MWENDA",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 27.0,
        "English": 30.0,
        "Kiswahili": 24.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "LIAM NJOROGE",
      "stream": "RED",
      "marks": {
        "Mathematics": 28.0,
        "English": 27.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "MALIK KIMUYA",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 28.0,
        "English": 30.0,
        "Kiswahili": 29.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "MARK ALVIN",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 25.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "MBUGUA GACHIRI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 26.0,
        "English": 30.0,
        "Kiswahili": 29.0,
        "Environmental Activities": 25.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "MICHAEL SAMUEL",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 26.0,
        "English": 28.0,
        "Kiswahili": 29.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "NATHAN EDWARD",
      "stream": "RED",
      "marks": {
        "Mathematics": null,
        "English": null,
        "Kiswahili": null,
        "Environmental Activities": null,
        "Christian Religious Education": null,
        "Creative Arts": null
      }
    },
    {
      "name": "RAYVON NG'ANG'A",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 29.0,
        "English": 29.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "REIGN MACHARIA",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 26.0,
        "English": 27.0,
        "Kiswahili": 20.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "RUBY KENA",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 29.0,
        "English": 30.0,
        "Kiswahili": 29.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "SAMARA MUMBI",
      "stream": "RED",
      "marks": {
        "Mathematics": 29.0,
        "English": 26.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "SARAH QUEEN",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 27.0,
        "English": 30.0,
        "Kiswahili": 28.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "SHANAYAH NYAMBURA",
      "stream": "RED",
      "marks": {
        "Mathematics": 28.0,
        "English": 26.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "SHANNEL KURIA",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 26.0,
        "English": null,
        "Kiswahili": 29.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "SOPHIA WACUKA",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 29.0,
        "English": 30.0,
        "Kiswahili": 29.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "TAMARA MAKENA",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 29.0,
        "English": 29.0,
        "Kiswahili": 29.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "TAMARA VANESSA",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 29.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "TEAIRRA WANJIKU",
      "stream": "RED",
      "marks": {
        "Mathematics": 28.0,
        "English": 30.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "THAYU GACHUNGI",
      "stream": "RED",
      "marks": {
        "Mathematics": 27.0,
        "English": 29.0,
        "Kiswahili": 28.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "YASMIN RAQUEL",
      "stream": "RED",
      "marks": {
        "Mathematics": 27.0,
        "English": 26.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 29.0
      }
    }
  ],
  "Grade 2": [
    {
      "name": "ABIGAEL WANGUI",
      "stream": "RED",
      "marks": {
        "Mathematics": 29.0,
        "English": 29.0,
        "Kiswahili": 28.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "AGATHA WEND0",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 30.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "ALVIN CHEGE",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 27.0,
        "English": 28.0,
        "Kiswahili": 24.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "AMARI MWINZI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 29.0,
        "English": 28.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "ARIANA ZURI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 27.0,
        "English": 27.0,
        "Kiswahili": 24.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "ARVIN NDARUA",
      "stream": "RED",
      "marks": {
        "Mathematics": 29.0,
        "English": 27.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "BLESSING MAKENA",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 20.0,
        "English": 26.0,
        "Kiswahili": 21.0,
        "Environmental Activities": 20.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 20.0
      }
    },
    {
      "name": "BONARERI OCHENGO",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 28.0,
        "English": 28.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "CARL KAMAU",
      "stream": "RED",
      "marks": {
        "Mathematics": 29.0,
        "English": 27.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "DAVID JABARI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 27.0,
        "English": 27.0,
        "Kiswahili": 26.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "ELLA HADASSAH",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 22.0,
        "English": 23.0,
        "Kiswahili": 20.0,
        "Environmental Activities": 21.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "ESTHER GATHONI",
      "stream": "RED",
      "marks": {
        "Mathematics": 28.0,
        "English": 28.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "ETHAN MURIMI",
      "stream": "RED",
      "marks": {
        "Mathematics": 29.0,
        "English": 29.0,
        "Kiswahili": 29.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "EUPHRATES TRIUMPH",
      "stream": "RED",
      "marks": {
        "Mathematics": 18.0,
        "English": 17.0,
        "Kiswahili": 10.0,
        "Environmental Activities": 21.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "GOLD OMONDI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 26.0,
        "English": 28.0,
        "Kiswahili": 21.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "HAZEL NJOKI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 28.0,
        "English": 29.0,
        "Kiswahili": 28.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "ISRAELLA WAITHERA",
      "stream": "RED",
      "marks": {
        "Mathematics": 18.0,
        "English": 25.0,
        "Kiswahili": 20.0,
        "Environmental Activities": 24.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 22.0
      }
    },
    {
      "name": "IVANNA WANJIRU",
      "stream": "RED",
      "marks": {
        "Mathematics": 23.0,
        "English": 26.0,
        "Kiswahili": 23.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "JAHEIM MUTUKU",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 28.0,
        "English": 26.0,
        "Kiswahili": 25.0,
        "Environmental Activities": 24.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 24.0
      }
    },
    {
      "name": "JAYLEN HARRY",
      "stream": "RED",
      "marks": {
        "Mathematics": 26.0,
        "English": 27.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "JESSE MUIGAI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 25.0,
        "English": 29.0,
        "Kiswahili": 24.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "JOANNA MURINGE",
      "stream": "RED",
      "marks": {
        "Mathematics": 29.0,
        "English": 27.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "JULIESTEVENS OFUYA",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 28.0,
        "English": 26.0,
        "Kiswahili": 26.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "JUSTIN GICHICHI",
      "stream": "RED",
      "marks": {
        "Mathematics": 25.0,
        "English": 27.0,
        "Kiswahili": 26.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "JYSON MUNENGE",
      "stream": "RED",
      "marks": {
        "Mathematics": 26.0,
        "English": 27.0,
        "Kiswahili": 24.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "KAMSI OKAFOR",
      "stream": "RED",
      "marks": {
        "Mathematics": 19.0,
        "English": 21.0,
        "Kiswahili": 22.0,
        "Environmental Activities": 24.0,
        "Christian Religious Education": 23.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "KHERI OREOLUWA",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 26.0,
        "English": 28.0,
        "Kiswahili": 28.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "KLOE WANGECI",
      "stream": "RED",
      "marks": {
        "Mathematics": 25.0,
        "English": 27.0,
        "Kiswahili": 24.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "LENNIX MWANGI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": null,
        "English": null,
        "Kiswahili": null,
        "Environmental Activities": null,
        "Christian Religious Education": null,
        "Creative Arts": null
      }
    },
    {
      "name": "LEO GITONGA",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": null,
        "English": null,
        "Kiswahili": null,
        "Environmental Activities": null,
        "Christian Religious Education": null,
        "Creative Arts": null
      }
    },
    {
      "name": "LORAINE KEMUNTO",
      "stream": "RED",
      "marks": {
        "Mathematics": 26.0,
        "English": 26.0,
        "Kiswahili": 25.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "MALCOM MUIGAI",
      "stream": "RED",
      "marks": {
        "Mathematics": 20.0,
        "English": 26.0,
        "Kiswahili": 24.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 22.0
      }
    },
    {
      "name": "MARK DAVID MAKORI",
      "stream": "RED",
      "marks": {
        "Mathematics": 27.0,
        "English": 28.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "MATHEW TUMUTI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 29.0,
        "English": 28.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "NATASHA ANGEL",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 26.0,
        "English": 28.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 24.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 24.0
      }
    },
    {
      "name": "NATHAN KIPLIMO",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 29.0,
        "English": 29.0,
        "Kiswahili": 24.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "NEEMA MUNGUTI",
      "stream": "RED",
      "marks": {
        "Mathematics": 26.0,
        "English": 27.0,
        "Kiswahili": 26.0,
        "Environmental Activities": 25.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "NYLA SIVANTOI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 28.0,
        "English": 28.0,
        "Kiswahili": 29.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "ORDELL MUNENE",
      "stream": "RED",
      "marks": {
        "Mathematics": 28.0,
        "English": 25.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 22.0
      }
    },
    {
      "name": "SALMAH GLORY",
      "stream": "RED",
      "marks": {
        "Mathematics": 29.0,
        "English": 29.0,
        "Kiswahili": 28.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "SHAWN NDIRANGU",
      "stream": "RED",
      "marks": {
        "Mathematics": 24.0,
        "English": 27.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "SKYE ABBY",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 23.0,
        "English": 28.0,
        "Kiswahili": 21.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "TAMALIA WAMUYU",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 24.0,
        "English": 28.0,
        "Kiswahili": 28.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "TAYLAN NJOROGE",
      "stream": "RED",
      "marks": {
        "Mathematics": 29.0,
        "English": 28.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "TEDDY MWANGI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 26.0,
        "English": 27.0,
        "Kiswahili": 15.0,
        "Environmental Activities": 20.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 20.0
      }
    },
    {
      "name": "TEHILAH ERIKA",
      "stream": "RED",
      "marks": {
        "Mathematics": 24.0,
        "English": 27.0,
        "Kiswahili": 24.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "TEHILAH WAITHIRA",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 27.0,
        "English": 28.0,
        "Kiswahili": 24.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "TRINITY MBAIRE",
      "stream": "RED",
      "marks": {
        "Mathematics": 28.0,
        "English": 27.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "VALERIE SHANGWE",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 29.0,
        "English": 27.0,
        "Kiswahili": 26.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "ZEMIRAH KATHURE",
      "stream": "RED",
      "marks": {
        "Mathematics": 27.0,
        "English": 28.0,
        "Kiswahili": 29.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 22.0
      }
    }
  ],
  "Grade 1": [
    {
      "name": "ADYNE TYCE",
      "stream": "RED",
      "marks": {
        "Mathematics": 29.0,
        "English": 27.0,
        "Kiswahili": 25.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "ALDEN KAMANGU",
      "stream": "RED",
      "marks": {
        "Mathematics": 28.0,
        "English": 27.0,
        "Kiswahili": 18.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "ANDY MUNDATI",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 29.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "ARIANNA WANGECI",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 23.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "BIANCAH NJERI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 27.0,
        "English": 24.0,
        "Kiswahili": 24.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "CHIRRIPO NORBETA",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 27.0,
        "English": 22.0,
        "Kiswahili": 18.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 24.0
      }
    },
    {
      "name": "CRECIA SWEENEY",
      "stream": "RED",
      "marks": {
        "Mathematics": 27.0,
        "English": 29.0,
        "Kiswahili": 25.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "DINAH WAKESHO",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 29.0,
        "Kiswahili": 24.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "ELLA CHEBET",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 23.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "ESME WANJIRU",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 29.0,
        "English": 29.0,
        "Kiswahili": 24.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "ETHAN MURIITHI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 29.0,
        "English": 30.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "GIANNA WANJIRU",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 29.0,
        "Kiswahili": 18.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "HANSEL NJAU",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 28.0,
        "English": 29.0,
        "Kiswahili": 28.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 24.0
      }
    },
    {
      "name": "HERI MUKENI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 25.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "JAYDEN BITI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 29.0,
        "English": 23.0,
        "Kiswahili": 18.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "JEREMY GACHUHI",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 28.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "JOSHUA BLESSING",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 29.0,
        "English": 30.0,
        "Kiswahili": 26.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "KAI KIPKOECH",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 24.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "LEILANI WANJIRU",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 26.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "LIAM NGIGE",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 23.0,
        "Kiswahili": 22.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "LIANA WANJIRU KIMANGI",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "LOGAN DANIEL TIROP",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 22.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "LUCIAS FAVOUR",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 25.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "MACHARIA TIFEOLUWA",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 26.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "MALEEK",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 28.0,
        "Kiswahili": 28.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "MYA WANJIRU",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 29.0,
        "English": 30.0,
        "Kiswahili": 23.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "NAILA NJOKI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 22.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "NIC WANJOHI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 30.0,
        "English": 29.0,
        "Kiswahili": 28.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "PAUL KINUTHIA",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "PETER KAIZEN",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 27.0,
        "English": 29.0,
        "Kiswahili": 19.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 24.0
      }
    },
    {
      "name": "PRINCE JAYDEN",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 30.0,
        "English": 27.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "RAPHAEL ALEXANDER",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 26.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "RIC BRANDON",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "RIO MATHENGE",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "ROY GACHERU",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 27.0,
        "English": 30.0,
        "Kiswahili": 24.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "RYAN NJEHU",
      "stream": "RED",
      "marks": {
        "Mathematics": 28.0,
        "English": 30.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 24.0
      }
    },
    {
      "name": "SASHA KENDI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 29.0,
        "English": 29.0,
        "Kiswahili": 25.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "SIOBHAN SHANY SETH",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 30.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "TALIA WAHITO",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 22.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "TEEJAY BARAKA",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 28.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "TEHILLAH WAIRIMU",
      "stream": "RED",
      "marks": {
        "Mathematics": 29.0,
        "English": 23.0,
        "Kiswahili": 18.0,
        "Environmental Activities": 24.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 24.0
      }
    },
    {
      "name": "THEO NESTON",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 26.0,
        "English": 25.0,
        "Kiswahili": 21.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 18.0
      }
    },
    {
      "name": "TORIA NDIRANGU",
      "stream": "RED",
      "marks": {
        "Mathematics": 27.0,
        "English": 27.0,
        "Kiswahili": 24.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "TYIANA ALYA",
      "stream": "RED",
      "marks": {
        "Mathematics": 28.0,
        "English": 30.0,
        "Kiswahili": 28.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 24.0
      }
    },
    {
      "name": "ZANE TAJI",
      "stream": "YELLOW",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 30.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "ZAYA ZAWADI",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 21.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "ZAYDAN NDICHU",
      "stream": "RED",
      "marks": {
        "Mathematics": 30.0,
        "English": 30.0,
        "Kiswahili": 27.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    }
  ],
  "PP2": [
    {
      "name": "Miguel Mwangi",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 29.0,
        "Language Activities": 28.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "Ellagrace Nyambura",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 29.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "Ayira Wanzila",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 30.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "Theo Gikonyo",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 29.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "Archie Ndarua",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 26.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "Katya Wambui",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 29.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "Grace Njeri",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 30.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "Mike Jayden",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 30.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "Tendai Aska Son",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 29.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "Laura Rose Atara",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 26.0,
        "Language Activities": 27.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "Lenny Ngigi",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 29.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "Eleanor Waithira",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 25.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 25.0
      }
    },
    {
      "name": "Gathura Gathura",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 29.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "Emmanuel Jesse",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": null,
        "Language Activities": null,
        "Environmental Activities": null,
        "Christian Religious Education": null,
        "Creative Arts": null
      }
    },
    {
      "name": "Cherub Ursla",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 29.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "Jude Jabari",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 25.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 25.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "Aidan Noah",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 30.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "Olivia Wanjiru Kemboi",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 30.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "Anna Makayler",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 28.0,
        "Language Activities": 26.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "KellySasha wanjiku",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 26.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "Hansel Wanjohi",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 20.0,
        "Language Activities": 20.0,
        "Environmental Activities": 20.0,
        "Christian Religious Education": 20.0,
        "Creative Arts": 20.0
      }
    },
    {
      "name": "Abiah Kyla Gacoya",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 25.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "Ace Odari",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 30.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "Brayden Mwendwa",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 28.0,
        "Language Activities": 27.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "Brianna Hadassah",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 29.0,
        "Language Activities": 30.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "Brielle Wanjiru",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 28.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "Darius Waweru Gatemba",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 24.0,
        "Language Activities": 25.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 22.0
      }
    },
    {
      "name": "Domani Shiwa Odhiambo",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 28.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "Dylan Kai",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 29.0,
        "Language Activities": 30.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "Elina Wanja",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 28.0,
        "Language Activities": 27.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "Gabriel Kiama",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 30.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "Hazel Muthoni",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 29.0,
        "Language Activities": 30.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "Ian Kingsley Kibe",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 28.0,
        "Language Activities": 30.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "Kaila Wael",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 28.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "Latifah Zawadi",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 29.0,
        "Language Activities": 30.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "Nathan Waiganjo",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 21.0,
        "Language Activities": 24.0,
        "Environmental Activities": 20.0,
        "Christian Religious Education": 20.0,
        "Creative Arts": 20.0
      }
    },
    {
      "name": "Nile Wanjama",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 21.0,
        "Language Activities": 25.0,
        "Environmental Activities": 25.0,
        "Christian Religious Education": 24.0,
        "Creative Arts": 20.0
      }
    },
    {
      "name": "Nova Waithera",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": null,
        "Language Activities": null,
        "Environmental Activities": null,
        "Christian Religious Education": null,
        "Creative Arts": null
      }
    },
    {
      "name": "Shanaya Muthoni",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 30.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "Shiloh Mali",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": null,
        "Language Activities": null,
        "Environmental Activities": null,
        "Christian Religious Education": null,
        "Creative Arts": null
      }
    },
    {
      "name": "Sky Muriu",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 25.0,
        "Language Activities": 30.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 29.0,
        "Creative Arts": 25.0
      }
    },
    {
      "name": "Taji Mbinda Maingi",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 28.0,
        "Language Activities": 28.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "Thayu  Matalanga",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 30.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "Zoey Grace Ngonyo",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 25.0,
        "Language Activities": 24.0,
        "Environmental Activities": 22.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 26.0
      }
    }
  ],
  "PP1": [
    {
      "name": "ALOYSIA  UTUGI",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 30.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "AYANNA  WAITHIRA",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 30.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "ACE NJOROGE",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 23.0,
        "Language Activities": 26.0,
        "Environmental Activities": 25.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 25.0
      }
    },
    {
      "name": "AMARI KINGI",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 29.0,
        "Language Activities": 28.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "BRIA NJERI",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 28.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "CHRISTIAN KARIUKI",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 27.0,
        "Language Activities": 24.0,
        "Environmental Activities": 18.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 24.0
      }
    },
    {
      "name": "CHRISTIAN NGANGA",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 29.0,
        "Language Activities": 28.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 24.0,
        "Creative Arts": 22.0
      }
    },
    {
      "name": "EZRA  MUNGAI",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 24.0,
        "Language Activities": 24.0,
        "Environmental Activities": 18.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 18.0
      }
    },
    {
      "name": "HAZEL KENDI",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 24.0,
        "Language Activities": 30.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "RIIRI YISRAEL",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 29.0,
        "Language Activities": 24.0,
        "Environmental Activities": 25.0,
        "Christian Religious Education": 25.0,
        "Creative Arts": 24.0
      }
    },
    {
      "name": "ISHMAEL JABULUM",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 22.0,
        "Language Activities": 18.0,
        "Environmental Activities": 15.0,
        "Christian Religious Education": 20.0,
        "Creative Arts": 15.0
      }
    },
    {
      "name": "JIANNA  WANJIRU",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 28.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "JENSEN TARAJI",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 30.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 24.0
      }
    },
    {
      "name": "JABALI TURING",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 30.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "JOSHUA OMONDI",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 28.0,
        "Language Activities": 30.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 25.0
      }
    },
    {
      "name": "KRYSTAL NJAMBI",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 25.0,
        "Environmental Activities": 25.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 24.0
      }
    },
    {
      "name": "LEYLANI KENDRA",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 30.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "LEON KARIUKI",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 28.0,
        "Language Activities": 28.0,
        "Environmental Activities": 25.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 25.0
      }
    },
    {
      "name": "MICHAELLA  NJERI",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 30.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 25.0
      }
    },
    {
      "name": "MAVERICK ETHAN",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 25.0,
        "Language Activities": 24.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 24.0
      }
    },
    {
      "name": "NASHLEY  WANJIKU",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 26.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "NAOMI  GACHIGI",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 26.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "NIMO WAMAITHA",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 14.0,
        "Language Activities": 10.0,
        "Environmental Activities": 20.0,
        "Christian Religious Education": 18.0,
        "Creative Arts": 15.0
      }
    },
    {
      "name": "NATHAN KIPCHUMBA",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 29.0,
        "Language Activities": 30.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 25.0
      }
    },
    {
      "name": "NADINE WANJIRU",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 28.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "NATE MACHARIA",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 30.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "NANDI MUTHONI",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 28.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 22.0
      }
    },
    {
      "name": "REAGAN MANDUKU",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 27.0,
        "Language Activities": 28.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 25.0
      }
    },
    {
      "name": "RAPHAEL CHEGE",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 28.0,
        "Language Activities": 28.0,
        "Environmental Activities": 24.0,
        "Christian Religious Education": 24.0,
        "Creative Arts": 24.0
      }
    },
    {
      "name": "TATIANA EMMA",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 23.0,
        "Language Activities": 22.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 24.0
      }
    },
    {
      "name": "TAJI KARUGA",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 28.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 30.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "THEO NDUATI",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 30.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "HELSA  NATALIE",
      "stream": "RED",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 28.0,
        "Environmental Activities": 18.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "Aydn Sanyaolu",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 30.0,
        "Environmental Activities": 30.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "Alph Karari",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 27.0,
        "Language Activities": 25.0,
        "Environmental Activities": null,
        "Christian Religious Education": null,
        "Creative Arts": null
      }
    },
    {
      "name": "Calvin Kariuki",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 18.0,
        "Language Activities": 25.0,
        "Environmental Activities": 24.0,
        "Christian Religious Education": 25.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "Cara Simantei",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 29.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "Ellie Wendo",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": null,
        "Language Activities": null,
        "Environmental Activities": null,
        "Christian Religious Education": null,
        "Creative Arts": null
      }
    },
    {
      "name": "Ethan Mwangi",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 28.0,
        "Language Activities": 24.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "Isla Makena",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 18.0,
        "Language Activities": 15.0,
        "Environmental Activities": 20.0,
        "Christian Religious Education": 18.0,
        "Creative Arts": 21.0
      }
    },
    {
      "name": "Ittai Amani",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 25.0,
        "Language Activities": 25.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "Ivannah Wambui",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 27.0,
        "Language Activities": 20.0,
        "Environmental Activities": 24.0,
        "Christian Religious Education": 22.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "Jeanette Njeri",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 29.0,
        "Language Activities": 30.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "Jevin Lbarunye",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 27.0,
        "Language Activities": 21.0,
        "Environmental Activities": 24.0,
        "Christian Religious Education": 20.0,
        "Creative Arts": 25.0
      }
    },
    {
      "name": "Kai Mayian",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 24.0,
        "Language Activities": 21.0,
        "Environmental Activities": 22.0,
        "Christian Religious Education": 23.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "Kalya Kipchumba",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 28.0,
        "Language Activities": 25.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "Kayla Wambui",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 28.0,
        "Language Activities": 28.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "Keren Jemma",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 29.0,
        "Environmental Activities": 29.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "Kylian Njomo",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 27.0,
        "Language Activities": 26.0,
        "Environmental Activities": 25.0,
        "Christian Religious Education": 25.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "Laurel Cherotich",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 26.0,
        "Language Activities": 23.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 25.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "Leikan Tenga",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 26.0,
        "Language Activities": 27.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 24.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "Lennox Chege",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": null,
        "Language Activities": null,
        "Environmental Activities": null,
        "Christian Religious Education": null,
        "Creative Arts": null
      }
    },
    {
      "name": "Leon Kamau",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 28.0,
        "Language Activities": 26.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 24.0,
        "Creative Arts": 25.0
      }
    },
    {
      "name": "Liam Gandi",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 14.0,
        "Language Activities": 14.0,
        "Environmental Activities": 20.0,
        "Christian Religious Education": 18.0,
        "Creative Arts": 18.0
      }
    },
    {
      "name": "Max Jelani",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 28.0,
        "Language Activities": 27.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "Maya Gathoni",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 27.0,
        "Language Activities": 26.0,
        "Environmental Activities": 27.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "Myles Ngugi",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 26.0,
        "Language Activities": 27.0,
        "Environmental Activities": 25.0,
        "Christian Religious Education": 25.0,
        "Creative Arts": 25.0
      }
    },
    {
      "name": "Nathan Thumbi",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 30.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 28.0,
        "Creative Arts": 30.0
      }
    },
    {
      "name": "Pendo Mmesoma",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 19.0,
        "Language Activities": 28.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "Precious Ndina",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 28.0,
        "Language Activities": 27.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "Rafa Kariuki",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 27.0,
        "Language Activities": 27.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 25.0,
        "Creative Arts": 27.0
      }
    },
    {
      "name": "Sabrina Wamuyu",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 30.0,
        "Language Activities": 29.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "Victoria Wangui",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 26.0,
        "Language Activities": 28.0,
        "Environmental Activities": 28.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 28.0
      }
    },
    {
      "name": "Winnie Wangui",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 28.0,
        "Language Activities": 26.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 27.0,
        "Creative Arts": 29.0
      }
    },
    {
      "name": "Zawadi Njoki",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 26.0,
        "Language Activities": 16.0,
        "Environmental Activities": 20.0,
        "Christian Religious Education": 18.0,
        "Creative Arts": 25.0
      }
    },
    {
      "name": "Zuri Njeri",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 25.0,
        "Language Activities": 26.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 26.0,
        "Creative Arts": 26.0
      }
    },
    {
      "name": "Zuri Wambui",
      "stream": "YELLOW",
      "marks": {
        "Mathematics Activities": 21.0,
        "Language Activities": 17.0,
        "Environmental Activities": 26.0,
        "Christian Religious Education": 25.0,
        "Creative Arts": 29.0
      }
    }
  ]
}

def title_case(name):
    return " ".join(w.capitalize() for w in str(name).split())

def run():
    app = create_app()
    with app.app_context():
        active_term = Term.query.filter_by(is_active=True).first()
        if not active_term:
            print("❌ No active term. Run seed.py first.")
            return

        # Get End Term assessment (number=3)
        end_term_ass = Assessment.query.filter_by(
            term_id=active_term.id, number=3
        ).first()
        if not end_term_ass:
            print("❌ End Term assessment not found.")
            return

        # Open it
        end_term_ass.is_open = True
        db.session.commit()
        print(f"✅ End Term assessment opened")

        total_saved = 0
        total_skipped = 0

        for grade_name, students in END_TERM_DATA.items():
            grade = Grade.query.filter_by(name=grade_name).first()
            if not grade:
                print(f"  ⚠️  Grade not found: {grade_name}")
                continue

            grade_saved = 0
            for s in students:
                name   = title_case(s["name"])
                stream_name = s.get("stream", "").upper()

                # Find student — match by name and grade
                student = Student.query.filter_by(
                    grade_id=grade.id
                ).filter(
                    Student.full_name.ilike(f"%{name.split()[0]}%")
                ).first()

                # Try exact match first
                exact = Student.query.filter_by(
                    full_name=name, grade_id=grade.id
                ).first()
                if exact:
                    student = exact

                if not student:
                    total_skipped += 1
                    continue

                marks = s.get("marks", {})
                for subject, val in marks.items():
                    if val is None:
                        continue

                    if isinstance(val, dict):
                        p1  = val.get("p1")
                        p2  = val.get("p2")
                        pct = val.get("pct")
                        if p1 is None and p2 is None:
                            continue
                        combined = float(pct) if pct else combine_split_subject(
                            p1, p2,
                            50 if grade_name in ["Grade 7","Grade 8","Grade 9"] else 40,
                            50 if grade_name in ["Grade 7","Grade 8","Grade 9"] else 10,
                        )
                        effective = combined
                        paper1, paper2 = p1, p2
                    else:
                        effective = float(val) if val is not None else None
                        paper1 = paper2 = combined = None
                        if effective is None:
                            continue

                    code, label = assign_performance_level(effective, grade_name)

                    mark = Mark.query.filter_by(
                        student_id=student.id,
                        assessment_id=end_term_ass.id,
                        subject=subject,
                    ).first()
                    if not mark:
                        mark = Mark(
                            student_id=student.id,
                            assessment_id=end_term_ass.id,
                            subject=subject,
                        )
                        db.session.add(mark)

                    if isinstance(val, dict):
                        mark.paper1_score   = paper1
                        mark.paper2_score   = paper2
                        mark.combined_score = combined
                        mark.score          = None
                    else:
                        mark.score          = effective
                        mark.paper1_score   = None
                        mark.paper2_score   = None
                        mark.combined_score = None

                    mark.grade_code  = code
                    mark.grade_label = label

                grade_saved += 1
                total_saved += 1

            db.session.commit()
            print(f"  ✅ {grade_name:10s} — {grade_saved} students saved")

        print(f"\n{'='*50}")
        print(f"  ✅ End Term import complete!")
        print(f"  Marks saved:   {total_saved}")
        print(f"  Skipped:       {total_skipped}")
        print(f"{'='*50}\n")

if __name__ == "__main__":
    run()
