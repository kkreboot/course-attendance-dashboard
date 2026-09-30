#!/usr/bin/env python3
"""Every invariant this project relies on, checked in one pass.

These are the same things a careful human verifies by hand after a change -
roll sets match the register, seat numbers run 1..N with no gaps, nobody
changed block by accident, the printed sheets actually describe the allocation
they claim to. Two real bugs in this project's history would have been caught
here without anyone going looking: an Overview counter reading 0 classes while
nine had been held, and one cohort's signature sheets silently rebuilt from
the *other* cohort's allocation (two blocks' PDFs deleted in the process).

Each check returns `ok` / `warn` / `fail`:
  fail - someone will hand out or act on something wrong
  warn - true but stale, or unverifiable from here
  ok   - verified, with the numbers that were checked
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import pandas as pd

import attendance_report as attrep
import config as C
import finder
import history
import mailer
import seating

OUT = Path("out")
ATT_WB = seating.ATTENDANCE_WORKBOOK
# Extra cohort-folder -> sheet names, for a folder whose name doesn't match its
# sheet. Normally empty: `attrep.sheet_for_cohort` pairs `out/english` with the
# `English` sheet by name.
COHORT_SHEET: dict[str, str] = {}


def sheet_for(cohort: str, path: str | Path | None = None) -> str | None:
    """The register sheet for cohort folder `cohort` in workbook `path`."""
    if cohort in COHORT_SHEET:
        return COHORT_SHEET[cohort]
    try:
        return attrep.sheet_for_cohort(str(path or ATT_WB), cohort)
    except Exception:                               # noqa: BLE001 - unreadable = unmapped
        return None


def _result(name: str, status: str, detail: str, fix: str = "", group: str = "") -> dict:
    return {"check": name, "status": status, "detail": detail, "fix": fix, "group": group}


def _cohorts() -> list[Path]:
    return sorted(p for p in OUT.glob("*") if p.is_dir() and (p / "allocation.csv").exists()
                  and not p.name.startswith("."))


def _room_of(cdir: Path) -> str | None:
    for q in cdir.glob("*_seating.xlsx"):
        stem = q.stem[: -len("_seating")]
        if stem in C.ROOMS:
            return stem
    return None


def _seat_no(seat: str) -> int | None:
    m = re.search(r"(\d+)$", str(seat))
    return int(m.group(1)) if m else None


# ─────────────────────────────── the checks ──────────────────────────────
def check_allocation(cdir: Path) -> list[dict]:
    g = cdir.name.title()
    alloc = pd.read_csv(cdir / "allocation.csv", dtype={"Roll": str})
    out = []

    dup = sorted(alloc.Roll[alloc.Roll.duplicated()])
    out.append(_result("No duplicate roll numbers", "ok" if not dup else "fail",
                       f"{len(alloc)} rows, all distinct" if not dup else
                       f"{len(dup)} repeated: {', '.join(dup[:6])}",
                       "Remove the duplicate rows from allocation.csv.", g))

    room = _room_of(cdir)
    if room is None:
        out.append(_result("Room identifiable", "warn",
                           "No `<room>_seating.xlsx` naming a room in config.ROOMS.",
                           "Rebuild the seating workbook so the room is unambiguous.", g))
        return out

    seats = seating.seat_map(room)
    known = set(seats.Seat)
    unknown = sorted(set(alloc.Seat) - known)
    out.append(_result("Seats exist in the room", "ok" if not unknown else "fail",
                       f"all {len(alloc)} seats are real seats in {room}" if not unknown else
                       f"{len(unknown)} not in {room}'s seat map: {', '.join(unknown[:6])}",
                       "Re-run the allocation, or move the affected students.", g))

    dup_seat = sorted(alloc.Seat[alloc.Seat.duplicated()])
    out.append(_result("No two students in one seat", "ok" if not dup_seat else "fail",
                       "every seat used at most once" if not dup_seat else
                       f"{len(dup_seat)} shared: {', '.join(map(str, dup_seat[:6]))}",
                       "Move one of each pair to a vacant seat.", g))

    bad_gap, bad_cap = [], []
    for block, sub in alloc.groupby("Block"):
        nums = sorted(n for n in (_seat_no(s) for s in sub.Seat) if n is not None)
        if nums != list(range(1, len(nums) + 1)):
            missing = [n for n in range(1, (max(nums) if nums else 0) + 1) if n not in nums]
            bad_gap.append(f"{block} (gap at {', '.join(map(str, missing[:3]))})")
        cap = len(seats[seats.Block == block])
        if len(sub) > cap:
            bad_cap.append(f"{block} ({len(sub)} in {cap} seats)")
    out.append(_result("Seat numbers contiguous from 1", "ok" if not bad_gap else "fail",
                       f"{alloc.Block.nunique()} blocks, no gaps" if not bad_gap else
                       "; ".join(bad_gap),
                       "Recompact the block - vacant capacity belongs at the end.", g))
    out.append(_result("Blocks within capacity", "ok" if not bad_cap else "fail",
                       "every block fits" if not bad_cap else "; ".join(bad_cap),
                       "Move the overflow to a block with room.", g))
    return out


def check_against_register(cdir: Path) -> list[dict]:
    """The Class Attendance workbook is the authority on who is in which batch
    and in what order; the allocation has to agree with it."""
    g = cdir.name.title()
    sheet = sheet_for(cdir.name) if Path(ATT_WB).exists() else None
    if sheet is None:
        return [_result("Matches the attendance register", "warn",
                        f"No workbook sheet mapped for `{cdir.name}`.",
                        "Name the folder after its sheet, or add it to checks.COHORT_SHEET.", g)]
    try:
        reg = seating.load_attendance_roster(ATT_WB, sheet)
    except (ValueError, KeyError, FileNotFoundError) as e:
        return [_result("Matches the attendance register", "fail", f"{type(e).__name__}: {e}",
                        "Check the workbook's header row and sheet names.", g)]

    alloc = pd.read_csv(cdir / "allocation.csv", dtype={"Roll": str})
    only_alloc = sorted(set(alloc.Roll) - set(reg.Roll))
    only_reg = sorted(set(reg.Roll) - set(alloc.Roll))
    out = [_result("Same students as the register", "ok" if not (only_alloc or only_reg) else "fail",
                   f"{len(alloc)} students, identical sets" if not (only_alloc or only_reg) else
                   f"seated but not in the register: {', '.join(only_alloc[:5]) or 'none'} · "
                   f"in the register but unseated: {', '.join(only_reg[:5]) or 'none'}",
                   "Seat the missing students, or remove the departed ones.", g)]

    rank = {str(r): i for i, r in enumerate(reg.Roll)}

    def _rank_of(r: str) -> float:
        if r in rank:
            return float(rank[r])
        b = seating.branch_code(r)
        branch_rolls = [ro for ro in reg.Roll if seating.branch_code(ro) == b]
        if branch_rolls:
            for prev in reversed(branch_rolls):
                if prev < r:
                    return rank[prev] + 0.5
            return rank[branch_rolls[0]] - 0.5
        return 10 ** 6

    wrong = []
    for block, sub in alloc.groupby("Block"):
        got = list(sub.sort_values(["SeatRow", "SeatCol"]).Roll)
        want = sorted(got, key=_rank_of)
        if got != want:
            wrong.append(block)
    out.append(_result("Sheets follow register order", "ok" if not wrong else "warn",
                       "every block reads in the workbook's own row order" if not wrong else
                       f"out of order: block(s) {', '.join(wrong)}",
                       "Re-sort with seating.reorder_within_blocks, then rebuild the sheets.", g))
    return out


def check_artifacts(cdir: Path) -> list[dict]:
    """Anything generated from `allocation.csv` must describe *this*
    allocation, and must not be older than it."""
    g = cdir.name.title()
    alloc_path = cdir / "allocation.csv"
    alloc = pd.read_csv(alloc_path, dtype={"Roll": str})
    a_mtime = alloc_path.stat().st_mtime
    out = []

    tpl_path = cdir / "sheets" / "template.json"
    if not tpl_path.exists():
        out.append(_result("Signature sheets built", "warn", "No `sheets/template.json` yet.",
                           "Build them on page 2.", g))
    else:
        tpl = json.loads(tpl_path.read_text())
        rows = {(r["roll"], r["seat"]) for b in tpl["blocks"].values() for r in b["rows"]}
        want = {(r.Roll, r.Seat) for r in alloc.itertuples()}
        missing, extra = want - rows, rows - want
        status = "ok" if not (missing or extra) else "fail"
        out.append(_result("Sheets match the allocation", status,
                           f"{len(rows)} printed rows, roll and seat identical" if status == "ok"
                           else f"{len(missing)} seated student(s) not on any sheet, "
                                f"{len(extra)} printed row(s) not in the allocation",
                           "Rebuild the sheets for THIS cohort - check the output folder.", g))

        room_tpl, room_dir = tpl.get("room"), _room_of(cdir)
        out.append(_result("Sheets name the right room",
                           "ok" if room_tpl == room_dir else "fail",
                           f"{room_tpl}" if room_tpl == room_dir else
                           f"sheets say {room_tpl!r}, this cohort's workbook says {room_dir!r}",
                           "Rebuild with the room the cohort actually sits in.", g))
        stale = tpl_path.stat().st_mtime < a_mtime - 1
        out.append(_result("Sheets newer than the allocation", "warn" if stale else "ok",
                           "sheets predate the current allocation.csv" if stale else "up to date",
                           "Rebuild the signature sheets.", g))

    for label, pattern in (("Seating workbook", "*_seating.xlsx"),
                           ("Hall plan", "*_seating_plan.html")):
        found = list(cdir.glob(pattern))
        if not found:
            out.append(_result(f"{label} built", "warn", f"No `{pattern}` in {cdir}.",
                               "Rebuild it on page 1.", g))
        else:
            stale = found[0].stat().st_mtime < a_mtime - 1
            out.append(_result(f"{label} current", "warn" if stale else "ok",
                               f"{found[0].name}" + (" predates allocation.csv" if stale else ""),
                               "Rebuild it.", g))
    return out


def check_published() -> list[dict]:
    """The two artifacts students actually see."""
    out = []
    allocs = {c.name: pd.read_csv(c / "allocation.csv", dtype={"Roll": str}) for c in _cohorts()}
    newest = max((c / "allocation.csv").stat().st_mtime for c in _cohorts()) if _cohorts() else 0

    page = Path("find-your-block.html")
    if not page.exists():
        out.append(_result("Find-your-block page", "warn", "Not built yet.",
                           "Build it on page 4.", "Published"))
    else:
        html = page.read_text(encoding="utf-8")
        missing, wrong = [], []
        for name, df in allocs.items():
            cdir = OUT / name
            room = _room_of(cdir)
            for s in df.itertuples():
                key = finder.roll_key(s.Roll)
                if f'"{key}"' not in html:
                    missing.append(s.Roll)
                elif room and f'"{key}":"{room}|{s.Block}"' not in html:
                    wrong.append(s.Roll)
        status = "fail" if (missing or wrong) else ("warn" if page.stat().st_mtime < newest - 1 else "ok")
        out.append(_result("Find-your-block page current", status,
                           f"{sum(len(d) for d in allocs.values())} students resolve correctly"
                           if status == "ok" else
                           (f"{len(missing)} missing, {len(wrong)} pointing at the wrong block"
                            if (missing or wrong) else "built before the current allocation"),
                           "Rebuild it on page 4.", "Published"))

    pdf = Path(f"{C.COURSE_CODE} Seating.pdf")
    if not pdf.exists():
        out.append(_result("Seat allotment PDF", "warn", "Not built yet.",
                           "Build it on page 3.", "Published"))
    else:
        stale = pdf.stat().st_mtime < newest - 1
        out.append(_result("Seat allotment PDF current", "warn" if stale else "ok",
                           "predates the current allocation" if stale else "up to date",
                           "Rebuild it on page 3.", "Published"))
    return out


def check_transitions_and_mail() -> list[dict]:
    out = []
    untracked = []
    for cdir in _cohorts():
        found = seating.detect_untracked_moves(cdir.name)
        if len(found):
            untracked.append(f"{cdir.name}: {len(found)}")
    baseline = [c.name for c in _cohorts() if not history.state_path(c.name).exists()]
    if baseline:
        out.append(_result("Hand-made moves detectable", "warn",
                           f"no baseline yet for {', '.join(baseline)}",
                           "Press “Adopt / re-baseline” on the Health check page.", "Transitions"))
    out.append(_result("No unrecorded block changes", "ok" if not untracked else "warn",
                       "every block change went through the toolkit" if not untracked
                       else "; ".join(untracked) + " changed outside it",
                       "Adopt them so the affected students can be mailed.", "Transitions"))

    pending = mailer.unmailed(seating.block_changed())
    out.append(_result("Everyone moved has been told", "ok" if pending.empty else "warn",
                       "nobody is waiting" if pending.empty else
                       f"{len(pending)} student(s) moved block and haven't been mailed",
                       "Mail them from the Email students page.", "Transitions"))
    return out


def check_attendance() -> list[dict]:
    out = []
    if not Path(ATT_WB).exists():
        return [_result("Attendance workbook present", "fail", f"`{ATT_WB}` not found.",
                        "Put the Institute's workbook back in the project folder.", "Attendance")]
    try:
        sheets = attrep.sheet_names(ATT_WB)
    except Exception as e:                              # noqa: BLE001 - reported verbatim
        return [_result("Attendance workbook readable", "fail", f"{type(e).__name__}: {e}",
                        "Check the file isn't open/locked or half-saved.", "Attendance")]

    for cdir in _cohorts():
        want = sheet_for(cdir.name)
        g = cdir.name.title()
        if want not in sheets:
            out.append(_result("Register sheet present", "fail",
                               f"no sheet matching {cdir.name!r} (found {', '.join(sheets)})",
                               "Rename the sheet, or update checks.COHORT_SHEET.", g))
            continue
        try:
            students, sessions = attrep.load(ATT_WB, want)
        except (ValueError, KeyError) as e:
            out.append(_result("Register columns detected", "fail", str(e),
                               "The header row (row 3) needs Roll Number / Name / %.", g))
            continue
        held = sessions[(sessions.Present + sessions.Absent + sessions.Excused) > 0]
        out.append(_result("Register readable", "ok",
                           f"{len(students)} students · {len(held)} of {len(sessions)} "
                           f"scheduled classes marked", "", g))

        # A cell is "marked" if it says Y, N, E or N/A. Held excludes excused
        # sessions under the default mode, so compare the raw mark count -
        # otherwise every student with an E looks like a data-entry gap.
        marked = (students.Present + students.Absent + students.Excused
                  + students.NotApplicable)
        odd = students[(marked > 0) & (marked < held.shape[0])]
        if len(odd):
            out.append(_result("Everyone marked for every class held", "warn",
                               f"{len(odd)} student(s) have blank cells in classes that were "
                               f"held (e.g. {', '.join(odd.Roll.head(4))})",
                               "A blank counts as neither present nor absent - mark it Y, N "
                               "or E (excused).", g))
        n_na = int(students.NotApplicable.sum())
        if n_na:
            out.append(_result("Not-applicable cells (N/A)", "ok",
                               f"{n_na} cell(s) across {int((students.NotApplicable > 0).sum())} "
                               f"student(s) - classes before a late admission joined",
                               "", g))
        n_exc = int(students.Excused.sum())
        if n_exc:
            mode = ("excluded from their percentage"
                    if attrep.EXCUSED_MODE == "exclude" else "counted as present")
            out.append(_result("Excused absences (E)", "ok",
                               f"{n_exc} excused mark(s) across "
                               f"{int((students.Excused > 0).sum())} student(s), {mode}",
                               "", g))
    return out


def check_combined_sheet() -> list[dict]:
    """The workbook's combined sheet is both batches on one list. It is a
    second copy of the same marks, so the only question worth asking is
    whether the two copies still agree: same students, same batch label, same
    attendance. A combined sheet quietly out of step with the registers is
    worse than not having one, because it is the one people read."""
    out = []
    g = "Attendance"
    if not Path(ATT_WB).exists():
        return out
    try:
        combined = attrep.combined_sheets(ATT_WB)
    except Exception as e:                              # noqa: BLE001 - reported verbatim
        return [_result("Combined sheet readable", "fail", f"{type(e).__name__}: {e}",
                        "Check the workbook isn't half-saved.", g)]
    if not combined:
        return out

    for sheet in combined:
        try:
            comb, _ = attrep.load(ATT_WB, sheet)
            batch_of = dict(zip(
                seating.load_attendance_roster(ATT_WB, sheet).Roll,
                seating.load_attendance_roster(ATT_WB, sheet).Medium))
        except (ValueError, KeyError) as e:
            out.append(_result(f"{sheet} sheet readable", "fail", str(e),
                               "The header row needs Roll Number / Name / Batch / %.", g))
            continue

        per_sheet, bad_batch, mismatched = {}, [], []
        for name in attrep.cohort_sheets(ATT_WB):
            stu, _ = attrep.load(ATT_WB, name)
            for r in stu.itertuples():
                per_sheet[str(r.Roll)] = (name, int(r.Present), int(r.Absent), int(r.Excused))

        here = {str(r.Roll): r for r in comb.itertuples()}
        missing = sorted(set(per_sheet) - set(here))
        extra = sorted(set(here) - set(per_sheet))
        for roll, (name, p, a, e) in per_sheet.items():
            row = here.get(roll)
            if row is None:
                continue
            if batch_of.get(roll, "") != name:
                bad_batch.append(f"{roll} says {batch_of.get(roll, '')!r}, is on {name}")
            if (int(row.Present), int(row.Absent), int(row.Excused)) != (p, a, e):
                mismatched.append(roll)

        detail = f"{len(here)} students, same marks as the batch sheets"
        status, fix = "ok", ""
        if missing or extra:
            status = "fail"
            detail = (f"{len(missing)} student(s) on a batch sheet but not on {sheet}"
                      + (f" (e.g. {', '.join(missing[:3])})" if missing else "")
                      + f"; {len(extra)} the other way round"
                      + (f" (e.g. {', '.join(extra[:3])})" if extra else ""))
            fix = f"Rebuild {sheet} from the two batch sheets."
        elif bad_batch:
            status = "fail"
            detail = f"{len(bad_batch)} wrong Batch label: {'; '.join(bad_batch[:3])}"
            fix = "Fix the Batch column, or move the student on the batch sheets."
        elif mismatched:
            status = "fail"
            detail = (f"{len(mismatched)} student(s) whose marks differ from their batch "
                      f"sheet (e.g. {', '.join(mismatched[:4])})")
            fix = f"Re-copy those rows into {sheet}; the batch sheets are the register."
        out.append(_result(f"{sheet} sheet agrees with the batch sheets", status, detail, fix, g))
    return out


def check_dashboard_sheet() -> list[dict]:
    """The workbook's own `Dashboard` sheet against the register it is built
    from. Excel caches formula results, so a sheet edited and saved without a
    recalculation can show yesterday's counts - and that sheet is what the
    course reads. A mismatch is a warning, not a failure: the register is the
    authority and nothing downstream reads the Dashboard's numbers."""
    g = "Attendance"
    if not Path(ATT_WB).exists():
        return []
    try:
        blocks = attrep.dashboard_blocks(ATT_WB)
    except Exception as e:                              # noqa: BLE001 - reported verbatim
        return [_result("Dashboard sheet readable", "warn", f"{type(e).__name__}: {e}",
                        "It is a convenience sheet; the register still rules.", g)]
    if not blocks or not blocks.get("summary"):
        return []

    sheet = attrep.combined_sheets(ATT_WB) or attrep.cohort_sheets(ATT_WB)
    students, sessions = attrep.load(ATT_WB, sheet[0])
    marked = sessions[(sessions.Present + sessions.Absent + sessions.Excused) > 0]
    settings = blocks.get("settings", {})
    threshold = float(settings.get("threshold (%)") or 75)
    total = int(float(settings.get("total classes (course)") or len(sessions)))
    remaining = max(0, total - len(marked))

    seen = students[students.Held > 0]
    below = seen[seen.Percent < threshold]
    mine = {
        "students": len(students),
        "below threshold": len(below),
        "critical (below 50%)": int((seen.Percent < 50).sum()),
        "can't reach threshold by end": sum(
            1 for r in below.itertuples()
            if not mailer.can_still_reach(int(r.Present), int(r.Absent), remaining, threshold)),
    }
    drift = []
    for key, ours in mine.items():
        row = blocks["summary"].get(key, {})
        theirs = row.get("All")
        if theirs is None:
            continue
        if int(float(theirs)) != int(ours):
            drift.append(f"{key}: sheet {int(float(theirs))}, register {ours}")
    held_sheet = settings.get("classes held")
    if held_sheet is not None and int(float(held_sheet)) != len(marked):
        drift.append(f"classes held: sheet {int(float(held_sheet))}, register {len(marked)}")

    return [_result("Dashboard sheet agrees with the register",
                    "warn" if drift else "ok",
                    "; ".join(drift) if drift else
                    f"{len(students)} students, {len(marked)} of {total} classes, "
                    f"{len(below)} below {attrep_pct(threshold)}%",
                    "Open the workbook and let Excel recalculate, or fix the formulas."
                    if drift else "", g)]


def attrep_pct(value) -> str:
    out = f"{float(value):.1f}"
    return out[:-2] if out.endswith(".0") else out


def check_course_settings() -> list[dict]:
    """The Course setup page's file: present, readable, and not the demo's."""
    import course_settings as cs

    g = "Course"
    if not cs.exists():
        return [_result("Course set up", "warn",
                        f"No `{cs.SETTINGS_FILE}` - running on the demo course ({C.COURSE}).",
                        "Fill in Maintenance › Course setup.", g)]
    if not cs.read_raw():
        return [_result("Course set up", "fail", f"`{cs.SETTINGS_FILE}` is unreadable.",
                        "Open Course setup and save again (a snapshot of the old file is kept).", g)]
    s = cs.load()
    out = []
    bad = cs.problems(s)
    out.append(_result("Course settings valid", "fail" if bad else "ok",
                       "; ".join(bad[:3]) if bad else f"{C.COURSE} · {C.SESSION}",
                       "Correct them on the Course setup page." if bad else "", g))
    left = cs.demo_leftovers(s)
    if left:
        out.append(_result("No demo values left", "warn", "still the demo's: " + ", ".join(left),
                           "Replace them on the Course setup page.", g))
    missing = [name for name in (C.ATTENDANCE_WORKBOOK,) if not Path(name).exists()]
    if missing:
        out.append(_result("Course workbooks in place", "warn",
                           "not found: " + ", ".join(missing),
                           "Upload it on Course setup › Files.", g))
    return out


def run_all() -> pd.DataFrame:
    rows: list[dict] = check_course_settings()
    cohorts = _cohorts()
    if not cohorts:
        rows.append(_result("Any allocation at all", "fail",
                            "No `out/<cohort>/allocation.csv` found.",
                            "Allocate seats on page 1.", "Setup"))
    for cdir in cohorts:
        rows += check_allocation(cdir)
        rows += check_against_register(cdir)
        rows += check_artifacts(cdir)
    if cohorts:
        rows += check_published()
        rows += check_transitions_and_mail()
    rows += check_attendance()
    rows += check_combined_sheet()
    rows += check_dashboard_sheet()
    return pd.DataFrame(rows, columns=["group", "check", "status", "detail", "fix"])


if __name__ == "__main__":
    df = run_all()
    icon = {"ok": "✓", "warn": "!", "fail": "✗"}
    for r in df.itertuples():
        print(f"{icon.get(r.status, '?')} [{r.group}] {r.check}: {r.detail}")
    bad = (df.status != "ok").sum()
    print(f"\n{len(df)} checks · {(df.status == 'ok').sum()} ok · "
          f"{(df.status == 'warn').sum()} warn · {(df.status == 'fail').sum()} fail")
    raise SystemExit(1 if (df.status == "fail").any() else 0)
