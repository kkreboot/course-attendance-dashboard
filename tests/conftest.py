"""Tests read the sample workbooks and saved exams from the project root, as
the dashboard does. On a fresh clone they are not there yet, so build them
once from the invented dataset. An existing workbook is never overwritten."""
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
os.chdir(ROOT)
sys.path.insert(0, str(ROOT))

if not (ROOT / "out" / "exams" / "quiz-1" / "allocation.csv").exists():
    subprocess.run([sys.executable, str(ROOT / "make_sample_data.py")], check=True, cwd=ROOT)
