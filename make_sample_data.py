#!/usr/bin/env python3
"""Build a complete, invented dataset so the dashboard runs from a fresh clone.

Nothing in here is real. Names are random pairs from two short lists, roll
numbers use the Institute's *format* (`B26CS1503`) with serials no real batch
uses, TAs are "TA 01" ... "TA 21", and every attendance mark is drawn from a
seeded random generator. The shapes, though, are the real ones: the same
header rows, the same session columns and N/A spacers, the same Combined and
Dashboard sheets, the same TA duty layout. That is what lets every parser,
page and test in the toolkit run unchanged against it.

    python make_sample_data.py                # installs sample_data/ into the project
    python make_sample_data.py --regenerate   # rebuild sample_data/ from the seed first
    python make_sample_data.py --no-install   # leave the project root alone

Installing copies the workbooks into the project root (where the dashboard
looks for them) and builds `out/`: both classroom allocations with their
signature sheets, and the Quiz 1 and Minor Exam seatings. It never
overwrites a workbook that is already there, so running it in a folder that
holds your real files is harmless; pass --force to replace them.
"""
from __future__ import annotations

import argparse
import math
import random
import shutil
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

import config as C

HERE = Path(__file__).resolve().parent
SAMPLE = HERE / "sample_data"
SEED = 1012

ATTENDANCE_WB = C.ATTENDANCE_WORKBOOK
TA_WB = C.TA_WORKBOOK
ROLL_LIST = C.ROLL_LIST
ROOMS_WB = C.EXAM_HALLS_WORKBOOK          # hall geometry only, shipped as-is

TODAY = datetime(2026, 9, 29)               # sessions up to here carry marks
THRESHOLD = 75.0
TOTAL_CLASSES = 39
BENCHMARK = 85.0

FIRST = ["Aarav", "Aditi", "Akash", "Ananya", "Arjun", "Bhavna", "Chetan", "Deepa",
         "Dev", "Diya", "Farhan", "Gauri", "Harsh", "Isha", "Jatin", "Kavya", "Kunal",
         "Lakshmi", "Manav", "Meera", "Nikhil", "Neha", "Om", "Pooja", "Pranav",
         "Rhea", "Rohan", "Sanya", "Siddharth", "Sneha", "Tanvi", "Uday", "Varun",
         "Vidya", "Yash", "Zoya", "Aman", "Kiran", "Ritika", "Tarun"]
LAST = ["Sharma", "Verma", "Iyer", "Nair", "Reddy", "Patel", "Singh", "Gupta", "Das",
        "Rao", "Mehta", "Joshi", "Kulkarni", "Chauhan", "Menon", "Bose", "Pillai",
        "Saxena", "Tiwari", "Yadav", "Kapoor", "Bhat", "Mishra", "Pandey"]

# B26 students per branch (English, Hindi). Sizes echo a first-year physics
# course of ~365; the split keeps every branch group inside its block.
BRANCHES = {"AE": (15, 10), "BB": (13, 10), "CH": (13, 9), "CI": (13, 16),
            "CM": (26, 19), "CS": (25, 18), "CY": (5, 3), "EC": (16, 12),
            "EE": (18, 13), "MA": (10, 7), "ME": (23, 17), "MT": (14, 10),
            "PH": (8, 5)}
# Backlog students (older intakes), English then Hindi. AI has no B26 batch.
BACKLOG = {"English": ["B22EE905", "B22MT936", "B23CM1929", "B24CS1911", "B24ME1907",
                       "B25CM1909", "B25MT1906", "B25CY1901", "B25CI1930", "B25EE1914"],
           "Hindi": ["B22AI914", "B23EC1912", "B24CH1903", "B25CS1920", "B25ME1911",
                     "B25PH1904"]}
LATE_ADMISSION = "B22PH903"     # joins at C8: N/A for the first seven classes
LATE_FROM = 8

# Column plan of the register after the "%" column: a session label, or None
# for an N/A spacer (holiday, exam day). Dates run Mon/Wed/Fri from 5 Aug.
LAYOUT = (["C"] * 9 + [None] + ["C"] * 3 + [None] + ["C"] * 5 + [None, None]
          + ["C"] * 7 + [None] + ["C"] * 11 + [None] * 4 + ["C"] * 4 + [None])
DATES = ["08-05", "08-07", "08-10", "08-12", "08-14", "08-17", "08-19", "08-21",
         "08-24", "08-26", "08-28", "08-31", "09-02", "09-04", "09-07", "09-09",
         "09-11", "09-12", "09-14", "09-16", "09-18", "09-21", "09-23", "09-24",
         "09-25", "09-26", "09-28", "09-30", "10-02", "10-05", "10-07", "10-09",
         "10-12", "10-14", "10-16", "10-19", "10-21", "10-23", "10-26", "10-28",
         "10-30", "11-02", "11-04", "11-06", "11-09", "11-11", "11-13", "11-16",
         "11-17"]
ROW1_NOTES = {9: "Holiday", 13: "Holiday", 19: "MINOR EXAM", 20: "MINOR EXAM",
              28: "Holiday", 40: "SEMESTER BREAK", 48: "FEEDBACK"}

BOLD = Font(bold=True)
HEAD_FILL = PatternFill("solid", fgColor="DDE7F3")


# ───────────────────────────── people ─────────────────────────────────────

def make_students(rng: random.Random) -> list[dict]:
    names = set()

    def name():
        while True:
            n = f"{rng.choice(FIRST)} {rng.choice(LAST)}"
            if n not in names:
                names.add(n)
                return n

    out = []
    for br, (en, hi) in BRANCHES.items():
        serials = rng.sample(range(501, 700), en + hi)
        for i, s in enumerate(sorted(serials)):
            out.append(dict(Roll=f"B26{br}1{s}", Name=name(),
                            Batch="English" if i < en else "Hindi"))
    for batch, rolls in BACKLOG.items():
        out += [dict(Roll=r, Name=name(), Batch=batch) for r in rolls]
    out.append(dict(Roll=LATE_ADMISSION, Name=name(), Batch="English"))
    return out


def session_columns() -> list[tuple[str, datetime]]:
    """(label, date) per register column after "%": "C7" or "N/A"."""
    cols, n = [], 0
    for kind, md in zip(LAYOUT, DATES):
        d = datetime(2026, int(md[:2]), int(md[3:]))
        if kind:
            n += 1
            cols.append((f"C{n}", d))
        else:
            cols.append(("N/A", d))
    return cols


def draw_marks(rng: random.Random, students: list[dict], cols) -> None:
    """Y/N/E per held session. Most students attend well; a tail does not,
    so the defaulter list, both mail templates and every band have members."""
    for s in students:
        u = rng.random()
        p = (rng.uniform(0.88, 1.0) if u < 0.80 else
             rng.uniform(0.68, 0.85) if u < 0.93 else rng.uniform(0.15, 0.55))
        marks = []
        for label, d in cols:
            if label == "N/A" or d > TODAY:
                marks.append(None)
            elif s["Roll"] == LATE_ADMISSION and int(label[1:]) < LATE_FROM:
                marks.append("N/A")
            elif rng.random() < p:
                marks.append("Y")
            else:
                marks.append("E" if rng.random() < 0.05 else "N")
        s["Marks"] = marks
        y, n = marks.count("Y"), marks.count("N")
        s["Y"], s["N"], s["E"] = y, n, marks.count("E")
        s["Pct"] = round(y / (y + n) * 100, 2) if y + n else None


# ───────────────────────────── workbooks ──────────────────────────────────

def _header(ws, cols, lead: list[str]) -> None:
    ws.cell(1, 1, C.COURSE_CODE).font = BOLD
    first = len(lead) + 1
    ws.cell(2, first - 1, "Attendance")
    for i, h in enumerate(lead, 1):
        c = ws.cell(3, i, h)
        c.font, c.fill = BOLD, HEAD_FILL
    for j, (label, d) in enumerate(cols):
        col = first + j
        ws.cell(2, col, d).number_format = "dd-mmm"
        c = ws.cell(3, col, label)
        c.font, c.fill = BOLD, HEAD_FILL
        if j in ROW1_NOTES:
            ws.cell(1, col, ROW1_NOTES[j])
    ws.freeze_panes = ws.cell(4, first)
    ws.column_dimensions["B"].width = 13
    ws.column_dimensions["C"].width = 24


def _row(ws, r: int, values) -> None:
    # explicit rows: freeze_panes touches row 4, so ws.append would skip it
    for c, v in enumerate(values, 1):
        if v is not None:
            ws.cell(r, c, v)


def write_attendance(students, cols, path: Path) -> None:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for batch in ("English", "Hindi"):
        ws = wb.create_sheet(batch)
        _header(ws, cols, ["SR", "Roll Number", "Name", "%"])
        rows = [s for s in students if s["Batch"] == batch]
        for i, s in enumerate(rows, 1):
            _row(ws, 3 + i, [i, s["Roll"], s["Name"], s["Pct"], *s["Marks"]])

    ws = wb.create_sheet("Combined")
    _header(ws, cols, ["SR", "Roll Number", "Name", "Batch", "%"])
    ws.cell(2, 1, "COMBINED - BOTH BATCHES, ROLL ORDER")
    ws.column_dimensions["D"].width = 10
    for i, s in enumerate(sorted(students, key=lambda s: s["Roll"]), 1):
        _row(ws, 3 + i, [i, s["Roll"], s["Name"], s["Batch"], s["Pct"], *s["Marks"]])

    write_dashboard(wb.create_sheet("Dashboard"), students, cols)
    wb.save(path)


def write_dashboard(ws, students, cols) -> None:
    """Values, not formulas: the real sheet is live formulas, but openpyxl
    cannot compute them and the toolkit reads cached values."""
    held = sum(1 for label, d in cols if label != "N/A" and d <= TODAY)
    remaining = TOTAL_CLASSES - held

    def stats(group):
        seen = [s for s in group if s["Pct"] is not None]
        below = [s for s in seen if s["Pct"] < THRESHOLD]
        cant = [s for s in below
                if (s["Y"] + remaining) / (s["Y"] + s["N"] + remaining) * 100 < THRESHOLD]
        avg = sum(s["Pct"] for s in seen) / len(seen) if seen else 0
        return [len(group), len(below), len(below) / len(group) if group else 0,
                round(avg, 2), sum(1 for s in seen if s["Pct"] < 50), len(cant)], below, cant

    groups = {"All": students,
              "English": [s for s in students if s["Batch"] == "English"],
              "Hindi": [s for s in students if s["Batch"] == "Hindi"]}
    figures = {k: stats(v) for k, v in groups.items()}

    ws.append([f"{C.COURSE_CODE} Attendance Dashboard"])
    ws.append(["Sample data: every figure here is computed from the invented register."])
    ws.append(["Settings", "Value"])
    ws.append(["Threshold (%)", THRESHOLD])
    ws.append(["Batch filter", "All"])
    ws.append(["Classes held", held])
    ws.append(["Total classes (course)", TOTAL_CLASSES])
    ws.append(["Class benchmark (%)", BENCHMARK])
    ws.append([])
    ws.append(["Summary", *groups])
    labels = ["Students", "Below threshold", "% below threshold", "Average attendance (%)",
              "Critical (below 50%)", "Can't reach threshold by end"]
    for i, label in enumerate(labels):
        ws.append([label, *(figures[k][0][i] for k in groups)])
    ws.append([])
    ws.append(["Attendance band", "Students"])
    pcts = [s["Pct"] for s in students if s["Pct"] is not None]
    for band, lo, hi in (("< 50%", 0, 50), ("50-75%", 50, 75), ("75-90%", 75, 90),
                         ("90-100%", 90, 101)):
        ws.append([band, sum(1 for p in pcts if lo <= p < hi)])
    ws.append([])
    _, below, cant = figures["All"]
    ws.append([f"Students below {THRESHOLD:.0f}% (both batches): {len(below)}   |   "
               f"Cannot reach {THRESHOLD:.0f}% even if they attend all remaining "
               f"{remaining} classes: {len(cant)}   |   Sorted lowest first"])
    ws.append(["SR", "Roll Number", "Name", "Batch", "Attendance %", "Present (Y)",
               "Absent (N)", "Excused (E)", "Classes needed to reach threshold"])
    order = {s["Roll"]: i for i, s in enumerate(sorted(students, key=lambda s: s["Roll"]), 1)}
    for s in sorted(below, key=lambda s: (s["Pct"], s["Roll"])):
        ws.append([order[s["Roll"]], s["Roll"], s["Name"], s["Batch"], s["Pct"], s["Y"],
                   s["N"], s["E"], max(0, 3 * s["N"] - s["Y"])])
    for row in (3, 10, 18):
        for c in ws[row]:
            c.font = BOLD
    ws.column_dimensions["A"].width = 30


def write_roll_list(students, path: Path) -> None:
    """The department's roll-list export, which (like the real one) has
    missed the late admission."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append([f"ROLL LIST OF  {C.COURSE_TITLE}({C.COURSE_CODE}, LTP : 3-0-0, CRD : 3)"])
    ws.append(["FACULTY NAME :  Dr. Instructor Two"])
    ws.append(["Serial No", "Roll No", "Name", "Secno", "Admission Category",
               "Subject Category", "Institute Email"])
    rows = sorted((s for s in students if s["Roll"] != LATE_ADMISSION), key=lambda s: s["Roll"])
    for i, s in enumerate(rows, 1):
        ws.append([i, s["Roll"], s["Name"], "N", "Regular (Full Time)", "IS",
                   f"{s['Roll'].lower()}@{C.EMAIL_DOMAIN}"])
    wb.save(path)


def write_ta_duty(students, path: Path) -> None:
    tas = [(f"P2{4 + i % 2}PH09{i:02d}", f"TA {i:02d}") for i in range(1, 22)]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TA Assigned Duties"
    ws.append([f"{C.COURSE} · TA duty assignment · {C.SESSION} (sample)"])
    ws.append([])
    ws.append(["Sr", "TA Roll", "TA Name", "Duty", None, "Notes"])
    duties = ([("Evaluation Head", None, "Minor and Major Exam Management"),
               ("Attendance Head", None, "Attendance, Seating and Quiz Management")]
              + [("Evaluation", None, None)] * 7)
    for day, slots in (("Monday", ("8AM-8:50 AM", "5PM-5:50PM")),
                       ("Wednesday", ("8AM-8:50 AM", "1PM-1:50PM")),
                       ("Friday", ("8AM-8:50 AM", "1PM-1:50PM"))):
        for slot in slots:
            duties += [("Attendance", day, slot)] * 2
    for i, ((roll, name), (duty, day, note)) in enumerate(zip(tas, duties), 1):
        ws.append([i, roll, name, duty, day, note])
    for c in ws[3]:
        c.font = BOLD

    write_showing_sheet(wb.create_sheet("Quiz 1 Sheets Showing"), students, tas)
    wb.save(path)


def write_showing_sheet(ws, students, tas) -> None:
    """One row per TA per slot, in the layout `answer_showing.read_schedule`
    parses: each TA shows one or two branches, backlog joined to its branch."""
    b26 = {}
    backlog = {}
    for s in students:
        br = s["Roll"][3:5]
        if s["Roll"].startswith("B26"):
            b26[br] = b26.get(br, 0) + 1
        else:
            backlog[br] = backlog.get(br, 0) + 1
    orphan = sum(n for br, n in backlog.items() if br not in b26)
    pairs = [["MA", "CY"], ["CH"], ["CM"], ["ME"], ["EC"], ["MT"], ["CS"], ["AE"], ["BB"],
             ["PH"], ["CI"], ["EE"]]

    def label(brs, with_orphans=False):
        parts = [f"{b} ({b26[b]})" for b in brs]
        joined = sum(backlog.get(b, 0) for b in brs)
        if joined:
            parts.append(f"Backlog ({joined})")
        if with_orphans and orphan:
            parts.append(f"AI Backlog ({orphan})")
        return " + ".join(parts), sum(b26[b] for b in brs) + joined + (orphan if with_orphans else 0)

    ws.append([f"{C.COURSE} · Quiz 1 · Answer Script Showing (sample)", None, None, None,
               datetime(2026, 9, 14)])
    ws.append(["Sr", "TA Roll", "TA Name", "Batch/Department", "Room No. & Time Slot",
               "No. of Students", "TA Name"])
    rooms = ["PH101", "PH102", "PH104", "PH105"]
    slots = ["6:30PM to 7:00PM", "7:15PM to 7:45PM"]
    ta_iter = iter(tas[2:])
    sr = 0
    for k, brs in enumerate(pairs):
        room, slot = rooms[k // 3], slots[(k % 3) // 2]
        text, n = label(brs, with_orphans=(brs == ["PH"]))
        sr += 1
        roll, name = next(ta_iter)
        ws.append([sr, roll, name, text, f"{room} {slot}", n,
                   f"{tas[0][1]} & {tas[1][1]}" if k % 3 == 0 else None])
    # a slot shared by two TAs carries one count: the sum, on its first row
    rows = list(ws.iter_rows(min_row=3, values_only=False))
    first_of = {}
    for r in rows:
        key = r[4].value
        if key in first_of:
            first_of[key][5].value += r[5].value
            r[5].value = None
            r[4].value = None
        else:
            first_of[key] = r


# ───────────────────────────── install ────────────────────────────────────

def build_samples(rng) -> list[dict]:
    SAMPLE.mkdir(exist_ok=True)
    students = make_students(rng)
    cols = session_columns()
    draw_marks(rng, students, cols)
    write_attendance(students, cols, SAMPLE / ATTENDANCE_WB)
    write_roll_list(students, SAMPLE / ROLL_LIST)
    write_ta_duty(students, SAMPLE / TA_WB)
    return students


def install(force: bool) -> None:
    for name in (ATTENDANCE_WB, TA_WB, ROLL_LIST, ROOMS_WB):
        dst = HERE / name
        if dst.exists() and not force:
            print(f"  kept     {name} (already present; --force to replace)")
            continue
        shutil.copy2(SAMPLE / name, dst)
        print(f"  copied   {name}")
    build_outputs(force)


def build_outputs(force: bool) -> None:
    """Classroom seating for both batches, and the two scheduled exams, built
    by the toolkit's own functions - the same calls the dashboard makes."""
    import os
    os.chdir(HERE)
    import exam_plan
    import exam_rooms
    import reports
    import seating
    import sheets

    for cohort, room, sheet in (("english", "LHC110", "English"),
                                ("hindi", "LHC2 101", "Hindi")):
        out = HERE / "out" / cohort
        if (out / "allocation.csv").exists() and not force:
            print(f"  kept     out/{cohort} (already built)")
            continue
        roster = seating.load_attendance_roster(ATTENDANCE_WB, sheet=sheet)
        alloc, seats = seating.allocate_grouped(roster, room, C.BRANCH_GROUPS[room])
        out.mkdir(parents=True, exist_ok=True)
        alloc.to_csv(out / "allocation.csv", index=False)
        label = f"{sheet} Batch"
        reports.build_workbook(alloc, seats, str(out / f"{room}_seating.xlsx"), room=room,
                               course=C.COURSE, cohort=label,
                               session=C.SESSION, source=ATTENDANCE_WB)
        reports.build_html(alloc, seats, str(out / f"{room}_seating_plan.html"), room=room,
                           course=C.COURSE, cohort=label, session=C.SESSION)
        sheets.build_all(alloc, str(out / "sheets"), course=C.COURSE, room=room,
                         session=C.SESSION)
        print(f"  built    out/{cohort}: {len(alloc)} students in {room}")

    rooms = exam_rooms.load_exam_rooms(ROOMS_WB)
    roster = exam_plan.exam_roster(ROLL_LIST, ATTENDANCE_WB)
    for key, order in (("quiz1", ("LHC 110", "LHC 105")), ("minor", ("LHC 105", "LHC 110"))):
        exam = exam_plan.scheduled(key)
        slug = exam_plan.slugify(exam["title"])
        if (exam_plan.EXAM_ROOT / slug / exam_plan.ALLOC_FILE).exists() and not force:
            print(f"  kept     out/exams/{slug} (already built)")
            continue
        seats, spare = [], []
        for rl in order:
            for bk in sorted(rooms[rl].blocks):
                seats += [(rl, bk, s) for s in rooms[rl].blocks[bk].seats("side")]
                spare += [(rl, bk, s) for s in rooms[rl].blocks[bk].overflow_seats("side")]
        short = len(roster) - len(seats)
        if 0 < short <= exam_rooms.OVERFLOW_LIMIT:
            seats += spare[:short]
        if len(roster) > len(seats):
            raise SystemExit(f"{exam['title']}: {len(seats)} seats for {len(roster)} students")
        alloc = roster[["Roll", "Name"]].copy()
        alloc["Room"] = [r for r, _, _ in seats[:len(alloc)]]
        alloc["Block"] = [b for _, b, _ in seats[:len(alloc)]]
        alloc["Seat"] = [s for _, _, s in seats[:len(alloc)]]
        venue = "; ".join(f"{rl} ({exam_rooms.venue_note(rl, sorted(set(alloc[alloc.Room == rl].Block)))})"
                          for rl in order)
        exam_plan.build_all(alloc, rooms, meta=dict(
            title=exam["title"], when=exam_plan.when_label(exam), venue=venue,
            course=C.COURSE, spaced="side"))
        print(f"  built    out/exams/{slug}: {len(alloc)} students")


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--no-install", action="store_true",
                   help="only write sample_data/, leave the project root alone")
    p.add_argument("--force", action="store_true",
                   help="replace workbooks and outputs that already exist")
    p.add_argument("--regenerate", action="store_true",
                   help="rebuild sample_data/ itself (it is committed, so normally "
                        "it is only copied)")
    a = p.parse_args()
    generated = [ATTENDANCE_WB, TA_WB, ROLL_LIST]
    if a.regenerate or not all((SAMPLE / n).exists() for n in generated):
        students = build_samples(random.Random(SEED))
        en = sum(s["Batch"] == "English" for s in students)
        print(f"sample_data/: {len(students)} invented students ({en} English, "
              f"{len(students) - en} Hindi)")
    else:
        print("sample_data/: using the committed files (--regenerate to rebuild)")
    if not a.no_install:
        install(a.force)


if __name__ == "__main__":
    main()
