"""Course setup: the settings file, the helpers behind the page, and the page's
save path end to end (in a scratch copy of the project, so the real folder is
never touched)."""
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

import pandas as pd
import pytest

import course_settings as cs
import course_setup

ROOT = Path(__file__).resolve().parent.parent


# ───────────────────────────── settings file ─────────────────────────────
def test_defaults_are_the_demo_and_valid():
    s = cs.load(ROOT / "tests" / "_absent.json")
    assert s["course"]["code"] == "DEMO101"
    assert cs.problems(s) == []
    assert set(cs.demo_leftovers(s)) == {"institute.email_domain", "ta.name", "course.code"}


def test_partial_file_merges_over_defaults(tmp_path):
    p = tmp_path / "s.json"
    p.write_text(json.dumps({"course": {"code": "PHL1010", "junk": 1}, "unknown": {"x": 1}}))
    s = cs.load(p)
    assert s["course"]["code"] == "PHL1010" and s["course"]["title"] == "Demo Course"
    assert "junk" not in s["course"] and "unknown" not in s


def test_halls_are_replaced_not_merged(tmp_path):
    p = tmp_path / "s.json"
    hall = [{"key": "A", "side": "", "rows": 2, "cols": 3, "wedge": [], "section": "front"}]
    p.write_text(json.dumps({"rooms": {"Hall 1": hall}, "branch_groups": {}}))
    s = cs.load(p)
    assert list(s["rooms"]) == ["Hall 1"] and s["branch_groups"] == {}


def test_a_broken_file_falls_back_to_the_demo(tmp_path):
    p = tmp_path / "s.json"
    p.write_text("{ half a file")
    assert cs.read_raw(p) == {} and cs.load(p)["course"]["code"] == "DEMO101"


def test_save_round_trips_atomically(tmp_path):
    p = tmp_path / "s.json"
    s = cs.load(p)
    s["course"]["title"] = "Mechanics"
    cs.save(s, p)
    assert cs.load(p)["course"]["title"] == "Mechanics"
    assert cs.load(p)["saved"] and not (tmp_path / "s.json.tmp").exists()


@pytest.mark.parametrize("patch,needle", [
    ({"course": {"code": "PHL 1010"}}, "letters and digits"),
    ({"institute": {"email_domain": "@nowhere"}}, "email domain"),
    ({"attendance": {"levels": [9, 5]}}, "increasing"),
    ({"attendance": {"threshold": 0}}, "between 0 and 100"),
    ({"instructors": []}, "at least one course instructor"),
    ({"course": {"start_date": "2026-12-01", "end_date": "2026-08-01"}}, "before its start"),
    ({"rooms": {"H": [{"key": "AA", "rows": 1, "cols": 1, "wedge": [], "section": "front"}]},
      "branch_groups": {}}, "one capital letter"),
    ({"branch_groups": {"Nowhere": {"A": ["CS"]}}}, "isn't set up"),
])
def test_problems_catch_what_would_break_a_page(patch, needle):
    s = cs._merge(cs.DEFAULTS, patch)
    assert any(needle in p for p in cs.problems(s)), cs.problems(s)


def test_people_and_lists_parse_the_way_people_type_them():
    assert cs.parse_people("Dr A, a@x.in\n\nDr B <b@x.in>\nDr C") == [
        {"name": "Dr A", "email": "a@x.in"}, {"name": "Dr B", "email": "b@x.in"},
        {"name": "Dr C", "email": ""}]
    assert cs.parse_int_list("3, 2 2,1") == [3, 2, 2, 1]
    assert cs.parse_codes("cs, ee  me") == ["CS", "EE", "ME"]


# ───────────────────────────── page helpers ─────────────────────────────
def test_hall_table_round_trips():
    s = cs.load(ROOT / "tests" / "_absent.json")
    rooms, errs = course_setup.rooms_from_table(course_setup.rooms_table(s))
    assert not errs and rooms == s["rooms"]
    assert course_setup.groups_from_table(course_setup.groups_table(s)) == s["branch_groups"]
    cap = course_setup.capacity(rooms).set_index("Hall")
    assert cap.loc["LHC2 101", "Total"] == 4 * 11 * 4


def test_hall_table_reports_bad_wedges_and_skips_blank_rows():
    df = pd.DataFrame([{"Hall": "H", "Block": "a", "Side": "", "Rows": 2, "Seats per row": 3,
                        "Wedge": "2, x", "Section": ""},
                       {"Hall": "", "Block": "", "Side": "", "Rows": None,
                        "Seats per row": None, "Wedge": "", "Section": ""}])
    rooms, errs = course_setup.rooms_from_table(df)
    assert list(rooms) == ["H"] and rooms["H"][0]["key"] == "A"
    assert rooms["H"][0]["section"] == "front" and errs


def test_uploads_are_checked_before_they_are_installed(tmp_path):
    wb = Path(sorted(ROOT.glob("Class Attendance DEMO101.xlsx"))[0]).read_bytes()
    assert course_setup.check_upload("attendance", wb, "x.xlsx") == ""
    assert "Couldn't read" in course_setup.check_upload("attendance", b"not a workbook", "x.xlsx")
    roll = (ROOT / "student_rollList-DEMO101.xlsx").read_bytes()
    assert "register sheet" in course_setup.check_upload("attendance", roll, "r.xlsx")
    dest = course_setup.install_file("attendance", wb, "PHL1010", tmp_path)
    assert dest.name == "Class Attendance PHL1010.xlsx" and dest.read_bytes() == wb


def test_code_change_renames_only_what_would_be_lost(tmp_path):
    (tmp_path / "Class Attendance OLD1.xlsx").write_bytes(b"a")
    (tmp_path / "OLD1_TA_Duty_Assignment.xlsx").write_bytes(b"t")
    (tmp_path / "NEW1_TA_Duty_Assignment.xlsx").write_bytes(b"already there")
    pairs = course_setup.renames_for_code("OLD1", "NEW1", tmp_path)
    assert [(a.name, b.name) for a, b in pairs] == [
        ("Class Attendance OLD1.xlsx", "Class Attendance NEW1.xlsx")]
    course_setup.apply_renames(pairs)
    assert (tmp_path / "Class Attendance NEW1.xlsx").read_bytes() == b"a"
    assert (tmp_path / "NEW1_TA_Duty_Assignment.xlsx").read_bytes() == b"already there"
    assert course_setup.renames_for_code("X", "X", tmp_path) == []


# ───────────────────── settings reach the toolkit ─────────────────────
SCRIPT = r"""
import json, sys
import config as C, mailer
print(json.dumps({"course": C.COURSE, "rooms": list(C.ROOMS), "domain": C.EMAIL_DOMAIN,
                  "levels": mailer.ATTENDANCE_LEVELS, "thr": mailer.ATTENDANCE_THRESHOLD,
                  "sig": mailer.ATTENDANCE_SIGNATURE, "cc": mailer.DEFAULT_ATTENDANCE_CC,
                  "wb": C.ATTENDANCE_WORKBOOK, "email": mailer.student_email("B26CS1001"),
                  "groups": C.BRANCH_GROUPS}))
"""


def test_settings_file_drives_config_and_mailer(tmp_path):
    s = cs._merge(cs.DEFAULTS, {
        "course": {"code": "PHL1010", "title": "Physics"},
        "institute": {"email_domain": "iitj.ac.in", "short": "IITJ"},
        "instructors": [{"name": "Dr P", "email": "p@iitj.ac.in"}],
        "ta": {"name": "Krishna", "name_local": ""},
        "attendance": {"threshold": 80, "levels": [4, 8]},
        "rooms": {"Hall 9": [{"key": "A", "side": "", "rows": 5, "cols": 5, "wedge": [],
                              "section": "front"}]},
        "branch_groups": {"Hall 9": {"A": ["CS"]}},
    })
    f = tmp_path / "s.json"
    cs.save(s, f)
    env = {**os.environ, "COURSE_SETTINGS_FILE": str(f)}
    out = subprocess.run([sys.executable, "-c", SCRIPT], cwd=ROOT, env=env,
                         capture_output=True, text=True, check=True).stdout
    got = json.loads(out)
    assert got["course"] == "PHL1010 Physics" and got["rooms"] == ["Hall 9"]
    assert got["levels"] == [[4, "L1"], [8, "L2"]] and got["thr"] == 80.0
    assert got["sig"].startswith("Krishna\nTeaching Assistant, PHL1010 Physics")
    assert got["cc"] == "p@iitj.ac.in" and got["email"] == "b26cs1001@iitj.ac.in"
    assert got["wb"] == "Class Attendance PHL1010.xlsx" and got["groups"] == {"Hall 9": {"A": ["CS"]}}


# ───────────────────────────── the page ─────────────────────────────
@pytest.fixture
def scratch_project(tmp_path):
    """A copy of the project with the sample workbooks, run from there."""
    dst = tmp_path / "proj"
    ignore = shutil.ignore_patterns(".git", ".venv*", "__pycache__", ".claude", "out")
    shutil.copytree(ROOT, dst, ignore=ignore)
    shutil.copytree(ROOT / "out", dst / "out", ignore=shutil.ignore_patterns(".history"))
    return dst


def test_setup_page_saves_renames_and_reloads(scratch_project, monkeypatch):
    from streamlit.testing.v1 import AppTest

    settings = scratch_project / "course_settings.json"
    monkeypatch.chdir(scratch_project)
    monkeypatch.setenv("COURSE_SETTINGS_FILE", str(settings))
    for m in course_setup.TOOLKIT_MODULES:          # re-read the env var
        if m in sys.modules:
            __import__("importlib").reload(sys.modules[m])
    try:
        at = AppTest.from_file(str(scratch_project / "dashboard.py"), default_timeout=180)
        at.run()
        assert at.session_state["phl_page"] == "Course setup", "first run opens on setup"
        at.text_input(key="setup_code").set_value("PHL1010")
        at.text_input(key="setup_domain").set_value("iitj.ac.in")
        at.text_input(key="setup_ta_name").set_value("Krishna")
        next(b for b in at.button if b.label == "Save course settings").click().run()
        assert not at.exception and not at.error, [e.value for e in at.error]
        assert json.loads(settings.read_text())["course"]["code"] == "PHL1010"
        assert (scratch_project / "Class Attendance PHL1010.xlsx").exists()
        assert not (scratch_project / "Class Attendance DEMO101.xlsx").exists()
        assert "PHL1010" in " ".join(m.value for m in at.sidebar.markdown)

        # a bad value is refused and nothing is written
        before = settings.read_text()
        at.text_input(key="setup_levels").set_value("9, 5")
        next(b for b in at.button if b.label == "Save course settings").click().run()
        assert at.error and "increasing" in at.error[0].value
        assert settings.read_text() == before
    finally:
        monkeypatch.undo()
        for m in course_setup.TOOLKIT_MODULES:      # back to the demo for later tests
            if m in sys.modules:
                __import__("importlib").reload(sys.modules[m])
