"""Bug fixes and the lookup / trend / forecast / bundle features.

Run against the invented sample data, like the rest of the suite.
"""
import io
import json
import os
import socket
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest

import attendance_report as attrep
import bundle
import checks
import lockfile
import lookup
import mailer
import seating

WB = seating.ATTENDANCE_WORKBOOK


# ───────────────────────────── lockfile ─────────────────────────────
def _dead_pid() -> int:
    pid = 999_999
    while True:
        try:
            os.kill(pid, 0)
        except ProcessLookupError:
            return pid
        except PermissionError:
            pass
        pid -= 1


@pytest.mark.skipif(os.name != "posix", reason="pid probe is POSIX-only")
def test_lock_ignores_a_dead_process_on_this_machine(tmp_path):
    lock = tmp_path / ".dashboard.lock"
    now = datetime.now().isoformat(timespec="seconds")
    lock.write_text(json.dumps([
        {"host": socket.gethostname(), "pid": _dead_pid(), "heartbeat": now},
        {"host": "other-machine", "pid": 1, "heartbeat": now},
    ]))
    others = lockfile.heartbeat(lock)
    assert [o["host"] for o in others] == ["other-machine"]


def test_lock_names_each_machine_once():
    text = lockfile.describe([{"host": "vm", "minutes": 2.3}, {"host": "vm", "minutes": 0.3},
                              {"host": "mac", "minutes": 1.0}])
    assert text.count("**vm**") == 1 and "0.3 min" in text and "2.3 min" not in text


# ─────────────────────── cohort -> register sheet ───────────────────────
def test_sheet_for_cohort_matches_by_name_not_a_hardcoded_table():
    assert attrep.sheet_for_cohort(WB, "english") == "English"
    assert attrep.sheet_for_cohort(WB, "HINDI") == "Hindi"
    assert attrep.sheet_for_cohort(WB, "combined") is None      # a view, not a cohort
    assert checks.sheet_for("english") == "English"


def test_cohort_sheet_override_still_wins(monkeypatch):
    monkeypatch.setitem(checks.COHORT_SHEET, "batch-a", "English")
    assert checks.sheet_for("batch-a") == "English"


def test_dashboard_sheet_check_has_no_hardcoded_cohorts():
    src = Path("checks.py").read_text(encoding="utf-8")
    assert '{"english": "English"' not in src


# ─────────────────────────── reorder ───────────────────────────
def test_register_reorder_puts_blocks_back_in_register_order():
    alloc = pd.read_csv("out/english/allocation.csv", dtype={"Roll": str})
    reg = list(seating.load_attendance_roster(WB, "English").Roll)
    room = "LHC110"
    shuffled = alloc.copy()
    for b in shuffled.Block.unique():          # scramble each block by roll order
        shuffled = seating.reorder_block_in_roll_order(shuffled, room, b, mode="roll")
    fixed = shuffled.copy()
    for b in fixed.Block.unique():
        fixed = seating.reorder_block_in_roll_order(fixed, room, b, mode="register",
                                                    reg_order=reg)
    rank = {r: i for i, r in enumerate(reg)}
    for _, sub in fixed.groupby("Block"):
        ranks = [rank[r] for r in sub.sort_values(["SeatRow", "SeatCol"]).Roll]
        assert ranks == sorted(ranks)
    # nobody changed block
    assert (fixed.set_index("Roll").Block.sort_index()
            == alloc.set_index("Roll").Block.sort_index()).all()


def test_register_reorder_without_an_order_reads_every_batch_sheet():
    """No room->sheet guess: a Hindi block reorders by the register with no
    reg_order passed, which only worked for LHC110/LHC2 101 before."""
    alloc = pd.read_csv("out/hindi/allocation.csv", dtype={"Roll": str})
    reg = list(seating.load_attendance_roster(WB, "Hindi").Roll)
    b = alloc.Block.iloc[0]
    a = seating.reorder_block_in_roll_order(alloc, "LHC2 101", b, mode="register")
    c = seating.reorder_block_in_roll_order(alloc, "LHC2 101", b, mode="register", reg_order=reg)
    assert list(a[a.Block == b].Roll) == list(c[c.Block == b].Roll)


# ───────────────────────── attendance arithmetic ─────────────────────────
def test_dashes_count_as_not_applicable():
    assert {"-", "–", "—"} <= attrep.NOT_APPLICABLE


@pytest.mark.parametrize("p,a", [(17, 5), (20, 0), (15, 5), (3, 1), (30, 10), (0, 0), (9, 4)])
def test_misses_allowed_is_the_exact_edge(p, a):
    t = 75.0
    k = mailer.misses_allowed(p, a, t)
    if p + a and p / (p + a) * 100 >= t:
        assert p / (p + a + k) * 100 >= t - 1e-9
        assert p / (p + a + k + 1) * 100 < t
    else:
        assert k == 0


def test_marks_agree_with_the_per_student_counts():
    students, _ = attrep.load(WB, "English")
    marks = attrep.load_marks(WB, "English")
    counts = marks.pivot_table(index="Roll", columns="Mark", aggfunc="size", fill_value=0)
    for r in students.head(40).itertuples():
        row = counts.loc[r.Roll]
        assert row.get("Y", 0) == r.Present and row.get("N", 0) == r.Absent
        assert row.get("E", 0) == r.Excused


def test_trend_is_per_batch_and_matches_session_totals():
    marks = attrep.load_marks(WB, "Combined")
    trend = attrep.batch_trend(marks)
    assert set(trend.Batch) == {"English", "Hindi"}
    _, sess = attrep.load(WB, "English")
    held = sess[(sess.Present + sess.Absent + sess.Excused) > 0]
    eng = trend[trend.Batch == "English"]
    assert len(eng) == len(held)
    assert list(eng.Present) == list(held.Present)
    assert eng.Cumulative.between(0, 100).all()


def test_at_risk_lists_only_students_above_the_line_who_would_drop():
    students, _ = attrep.load(WB, "Combined")
    risk = attrep.at_risk(students, 75, 3)
    assert len(risk) and (risk.Percent >= 75).all() and (risk.CanMiss < 3).all()
    assert (risk.IfMissed < 75).all()
    # a horizon capped by the classes left: none left, nobody at risk
    assert attrep.at_risk(students, 75, 3, remaining=0).empty
    assert len(attrep.at_risk(students, 75, 3, remaining=1)) <= len(risk)
    assert attrep.at_risk(students, 75, 0).empty


# ─────────────────────────── student lookup ───────────────────────────
def test_search_finds_by_roll_and_by_name_words():
    people = lookup.directory()
    assert len(people) == people.Roll.nunique() >= 365
    roll = people.Roll.iloc[10]
    assert lookup.search(roll.lower(), people).Roll.iloc[0] == roll
    name = people.Name.iloc[10]
    assert roll in set(lookup.search(name.upper(), people).Roll)
    assert lookup.search("", people).empty


def test_record_gathers_seat_exams_and_attendance():
    alloc = pd.read_csv("out/english/allocation.csv", dtype={"Roll": str})
    roll = alloc.Roll.iloc[0]
    rec = lookup.student_record(roll)
    assert rec.found and rec.batch == "English"
    assert rec.classroom[0]["Seat"] == alloc.Seat.iloc[0]
    assert rec.exams, "sample exams seat everybody"
    students, _ = attrep.load(WB, "English")
    me = students[students.Roll == roll].iloc[0]
    assert rec.attendance["present"] == me.Present
    assert rec.attendance["absent"] == me.Absent
    assert len(rec.marks) and set(rec.marks.Mark) <= {"Y", "N", "E", "NA", ""}


def test_unknown_student_is_not_found():
    assert not lookup.student_record("ZZ99ZZ999").found


# ─────────────────────────── print bundle ───────────────────────────
def test_cohort_bundle_holds_the_sheets_and_a_manifest():
    data, n = bundle.cohort_bundle("out/english")
    names = zipfile.ZipFile(io.BytesIO(data)).namelist()
    assert "MANIFEST.txt" in names and n == len(names) - 1
    assert any(x.startswith("english/sheets/signature_block_") for x in names)
    assert not any("allocation.csv" in x or "template.json" in x for x in names)


def test_exam_bundle_and_stale_marking(tmp_path):
    exam = tmp_path / "quiz"
    exam.mkdir()
    (exam / "seat_map.html").write_text("<html></html>")
    os.utime(exam / "seat_map.html", (1_000_000, 1_000_000))
    (exam / "allocation.csv").write_text("Roll\nX\n")
    files = bundle.exam_files(exam)
    assert [a for _, a in files] == ["quiz/seat_map.html"]
    assert bundle.stale(files, exam / "allocation.csv") == ["quiz/seat_map.html"]
    assert "OLDER THAN ALLOCATION" in bundle.manifest(files, "t", exam / "allocation.csv")
