"""Tests read the sample workbooks and saved exams from the project root, as
the dashboard does. On a fresh clone they are not there yet, so build them
once from the invented dataset. An existing workbook is never overwritten."""
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
os.chdir(ROOT)
# The suite asserts on the demo course (DEMO101, example.edu, notices at 5/9/13
# absences). Point the settings at a file that doesn't exist, so a folder that
# has been through Course setup still tests against the demo values.
os.environ["COURSE_SETTINGS_FILE"] = str(ROOT / "tests" / "_no_course_settings.json")
sys.path.insert(0, str(ROOT))

if not (ROOT / "out" / "exams" / "quiz-1" / "allocation.csv").exists():
    subprocess.run([sys.executable, str(ROOT / "make_sample_data.py")], check=True, cwd=ROOT)
