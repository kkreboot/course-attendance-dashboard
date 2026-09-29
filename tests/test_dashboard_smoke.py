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
         "Attendance summary", "Email students", "Health check"]


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
