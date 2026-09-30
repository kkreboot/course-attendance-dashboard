"""Every page must render without raising.

`dashboard.py` is one long script; a typo in a page nobody opened that day
surfaces as a red traceback in front of whoever finally opens it. Streamlit's
own AppTest harness runs the script headlessly with a chosen page selected,
which catches exactly that.
"""
from pathlib import Path

import pytest

from streamlit.testing.v1 import AppTest

# AppTest resolves relative paths against *this* file, not the project root.
DASHBOARD = str(Path(__file__).resolve().parent.parent / "dashboard.py")

PAGES = ["Overview", "Allocate seats", "Signature sheets", "Seat allotment PDF",
         "Find your block", "Move a student", "Posters & packs", "Exam seating",
         "Answer script showing", "Student lookup", "Attendance summary",
         "Email students", "Health check", "Course setup"]


def run_page(name: str, **state) -> AppTest:
    at = AppTest.from_file(DASHBOARD, default_timeout=90)
    at.session_state["phl_page"] = name
    for k, v in state.items():
        at.session_state[k] = v
    at.run()
    return at


@pytest.mark.parametrize("page", PAGES)
def test_page_renders(page):
    at = run_page(page)
    assert not at.exception, f"{page} raised: {[e.value for e in at.exception]}"


def test_exam_page_offers_the_new_controls():
    at = run_page("Exam seating")
    labels = [r.label for r in at.radio]
    assert any("Seating order" in l for l in labels), "branch-mixing control missing"
    assert any("Seat spacing" in l for l in labels)


def test_health_check_runs_and_finds_no_failures():
    at = run_page("Health check")
    text = " ".join(m.value for m in at.markdown) + " ".join(str(m.value) for m in at.metric)
    assert not at.exception
    # the metric row is Passing / Warnings / Failures
    failures = [m for m in at.metric if m.label == "Failures"]
    assert failures and failures[0].value == "0", "health check reports a failure"


def test_every_sidebar_page_is_smoke_tested():
    import re
    src = (Path(DASHBOARD)).read_text(encoding="utf-8")
    block = src[src.index("PAGES = {"):src.index("PAGE_META")]
    names = re.findall(r'\("[a-z_]+", "([^"]+)", "', block)
    assert set(names) == set(PAGES), set(names) ^ set(PAGES)


def test_overview_lists_batches_not_the_exams_folder():
    at = run_page("Overview")
    labels = [e.label for e in at.expander]
    assert not any("out/exams" in l for l in labels), labels
    assert "bundle_pick" in [s.key for s in at.selectbox]


def test_attendance_report_pdf_builds_from_the_dashboard(tmp_path):
    """The build ran after the workbook's temporary copy was deleted, so this
    button always failed with FileNotFoundError."""
    out = tmp_path / "report.pdf"
    at = run_page("Attendance summary")
    at.text_input(key="attrep_pdf_name").set_value(str(out)).run()
    at.button(key="attrep_pdf_run").click().run()
    assert not at.exception and not at.error, [e.value for e in at.error]
    assert out.exists() and out.read_bytes()[:4] == b"%PDF"


def test_attendance_page_shows_trend_and_forecast():
    at = run_page("Attendance summary")
    heads = " ".join(m.value for m in at.markdown)
    assert "Attendance over the term" in heads and "At risk" in heads
    assert any(m.label == "At risk" for m in at.metric)


def test_student_lookup_finds_a_student():
    import pandas as pd
    roll = pd.read_csv("out/hindi/allocation.csv", dtype={"Roll": str}).Roll.iloc[0]
    at = run_page("Student lookup")
    at.text_input(key="lookup_query").set_value(roll.lower()).run()
    assert not at.exception
    assert any(m.label == "Classroom seat" and m.value.startswith("Block") for m in at.metric)
    assert any(roll in c.value for c in at.code)


def test_reorder_defaults_to_register_order():
    at = run_page("Move a student")
    opts = at.selectbox(key="reorder_mode_pick").options
    assert "register" in opts[0].lower()
