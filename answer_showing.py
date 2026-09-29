"""Quiz 1 answer script showing: one marks sheet per TA per time slot.

Each sheet lists the roll numbers and names of the students whose scripts a
TA shows in one room and time slot, with blank Marks / Tick / Remarks boxes
to fill in by pen before the scripts go to the students. A cover page carries
the full schedule: room, time slot, students per slot and the Evaluation TAs
on standby in each room.

The schedule is read from the "Quiz 1 Sheets Showing" tab of the TA duty
workbook, so a change there is a rerun here. Its shape: one row per TA per
slot, "Room No. & Time Slot" like "PH101 6:30PM to 7:00PM", No. of Students
merged down each slot's rows, and the Evaluation TAs ("A & B") merged down
each room's rows. A TA may appear on more than one row.

Students come from the Class Attendance workbook (the primary roster) and are
split by roll prefix: B26XX1234 belongs to branch XX; anything older than the
B26 intake is a backlog student and goes to the B22-B25 group.

    python answer_showing.py          # writes out/exams/quiz-1/answer_showing/
"""
from __future__ import annotations
import re
from dataclasses import dataclass
from pathlib import Path
import openpyxl
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
import config as C
import seating

TITLE = f"{C.COURSE} · Quiz 1 · Answer Script Showing"
ROSTER = C.ATTENDANCE_WORKBOOK
DUTY_BOOK = C.TA_WORKBOOK
DUTY_SHEET = "Quiz 1 Sheets Showing"
OUT = Path("out/exams/quiz-1/answer_showing")

EVAL_NOTE = ("The Evaluation TAs will also be available throughout the sheet-viewing "
             "process. They will be present to handle any issues or queries that the "
             "assigned TAs may face while showing the answer sheets.")


@dataclass
class Duty:
    roll: str
    name: str
    label: str
    groups: list[tuple[str, int]]
    room: str
    slot: str
    slot_students: int          # as written on the tab, for the whole slot
    evaluators: list[str]
    backlog: int = 0            # "Backlog (n)" on the tab: joined to this TA's B26 groups


def parse_label(label: str) -> tuple[list[tuple[str, int]], int]:
    """"CM (45) + Backlog (3)" -> ([("B26CM", 45)], 3). A bare department code
    is that department's B26 batch; "Backlog (n)" is backlog joined to it; "AI
    Backlog (n)" is backlog whose department has no B26 batch (group B22-B25).
    Full batch names ("B26CM (45)", "B22-B25 (17)") still read."""
    groups, backlog = [], 0
    for code, n in re.findall(r"\b([A-Z]{2})\s+Backlog\s*\((\d+)\)", label):
        groups.append(("B22-B25", int(n)))
    label = re.sub(r"\b[A-Z]{2}\s+Backlog\s*\(\d+\)", "", label)
    backlog = sum(int(n) for n in re.findall(r"Backlog\s*\((\d+)\)", label, re.I))
    label = re.sub(r"Backlog\s*\(\d+\)", "", label, flags=re.I)
    for name, n in re.findall(r"\b(B\d\d[\w-]*|[A-Z]{2})\s*\((\d+)\)", label):
        groups.append((name if name.startswith("B") and name[1:3].isdigit() else "B26" + name,
                       int(n)))
    return groups, backlog


def read_schedule(path: str = DUTY_BOOK) -> tuple[list[Duty], str]:
    """Duties in sheet order, and the date. Blank merged cells carry the value
    above: the slot count down its slot, the evaluators down their room."""
    ws = openpyxl.load_workbook(path, data_only=True)[DUTY_SHEET]
    rows = list(ws.iter_rows(max_col=7, values_only=True))
    date = rows[0][4]
    date = date.strftime("%-d %b %Y") if hasattr(date, "strftime") else str(date or "")
    hdr = next(i for i, r in enumerate(rows) if r[0] == "Sr")
    duties = []
    where = count = None
    evals: dict[str, list[str]] = {}
    for r in rows[hdr + 1:]:
        if not isinstance(r[0], (int, float)):
            break
        if r[4]:
            if str(r[4]).strip() != where:
                count = None
            where = str(r[4]).strip()
        if where is None:
            raise SystemExit(f"{DUTY_SHEET}: TA {r[2]} has no room")
        room, _, slot = where.partition(" ")
        if r[5]:
            count = int(r[5])
        if r[6]:
            evals[room] = [n.strip() for n in re.split(r"&|,|\band\b", str(r[6])) if n.strip()]
        label = str(r[3]).strip()
        groups, backlog = parse_label(label)
        duties.append(Duty(str(r[1]).strip(), str(r[2]).strip(), label, groups,
                           room, slot.strip(), count or 0, [], backlog))
    for d in duties:
        d.evaluators = evals.get(d.room, [])
    return duties, date


def load_groups(roster: str = ROSTER) -> dict[str, pd.DataFrame]:
    """Branch groups with a Backlog flag. A backlog student (roll before the
    B26 intake) joins the B26 group of their own department, listed after
    the B26 students; one whose department has no B26 batch stays in B22-B25."""
    r = seating.load_attendance_roster(roster)
    rolls = r.Roll.astype(str)
    r["Backlog"] = ~rolls.str.startswith("B26")
    b26 = {x[:5] for x in rolls[~r.Backlog]}
    r["Group"] = [x[:5] if not bl else ("B26" + x[3:5] if "B26" + x[3:5] in b26 else "B22-B25")
                  for x, bl in zip(rolls, r.Backlog)]
    return {g: d.sort_values(["Backlog", "Roll"]).reset_index(drop=True)
            for g, d in r.groupby("Group")}


def check(duties: list[Duty], groups: dict[str, pd.DataFrame]) -> None:
    """Every student shown exactly once, and the tab's counts true."""
    problems = []
    assigned = [g for d in duties for g, _ in d.groups]
    dup = sorted({g for g in assigned if assigned.count(g) > 1})
    missing = sorted(set(groups) - set(assigned))
    unknown = sorted(set(assigned) - set(groups))
    if dup or missing or unknown:
        problems.append(f"duplicated {dup}, unassigned {missing}, not on roster {unknown}")
    for d in duties:
        # "B26EE (31) + Backlog (2)": the B26 count, then the backlog joined to it.
        for g, n in d.groups:
            if g not in groups:
                continue
            real = len(groups[g]) if g == "B22-B25" else int((~groups[g].Backlog).sum())
            if real != n:
                problems.append(f"{d.name}: tab says {g} ({n}), roster has {real}")
        joined = sum(int(groups[g].Backlog.sum()) for g, _ in d.groups
                     if g in groups and g != "B22-B25")
        if joined != d.backlog:
            problems.append(f"{d.name}: tab says Backlog ({d.backlog}), roster has {joined}")
    for key in dict.fromkeys((d.room, d.slot) for d in duties):
        ds = [d for d in duties if (d.room, d.slot) == key]
        real = sum(len(groups.get(g, [])) for d in ds for g, _ in d.groups)
        if ds[0].slot_students and ds[0].slot_students != real:
            problems.append(f"{key[0]} {key[1]}: tab says {ds[0].slot_students} students, "
                            f"roster has {real}")
    if problems:
        raise SystemExit("Schedule does not match the roster:\n  " + "\n  ".join(problems))


def students_of(d: Duty, groups) -> pd.DataFrame:
    """B26 students of every group first, then all backlog students together."""
    s = pd.concat([groups[g] for g, _ in d.groups], ignore_index=True)
    return pd.concat([s[~s.Backlog], s[s.Backlog]], ignore_index=True)


BACKLOG_HEAD = "Backlog students"


def sheet_rows(d: Duty, groups) -> list[list[str]]:
    """Numbered student rows, with a heading row (no number) above the
    backlog students when they follow B26 students."""
    s = students_of(d, groups)
    rows, sr = [], 0
    for k, st in enumerate(s.itertuples()):
        if st.Backlog and k and not s.Backlog.iloc[k - 1]:
            rows.append([BACKLOG_HEAD, "", "", "", "", ""])
        sr += 1
        rows.append([str(sr), st.Roll, st.Name, "", "", ""])
    return rows


def _wrap(cv, text: str, width: float, font: str, size: float) -> list[str]:
    lines, cur = [], ""
    for w in text.split():
        t = f"{cur} {w}".strip()
        if cv.stringWidth(t, font, size) <= width:
            cur = t
        else:
            lines.append(cur)
            cur = w
    return lines + [cur] if cur else lines


W, H = A4
MX = 14 * mm


def _header(cv, subtitle: str, date: str) -> float:
    cv.setFont("Helvetica-Bold", 13)
    cv.drawString(MX, H - 16 * mm, TITLE)
    cv.drawRightString(W - MX, H - 16 * mm, date)
    cv.setFont("Helvetica", 8)
    cv.drawString(MX, H - 21 * mm, subtitle)
    cv.drawRightString(W - MX, H - 21 * mm, C.instructor_names())
    cv.setStrokeGray(0.3)
    cv.setLineWidth(0.8)
    cv.line(MX, H - 23.5 * mm, W - MX, H - 23.5 * mm)
    return H - 23.5 * mm


def _table(cv, top: float, cols: list[tuple[str, float]], rows: list[list[str]],
           row_h: float, size: float = 9, boxes=frozenset(), breaks=None,
           shade: bool = True, body_font: str = "Helvetica-Bold") -> float:
    """`boxes`: column edges to rule; `breaks`: {row index: line width} for a
    heavier rule above that row (a new room or slot on the cover)."""
    breaks = breaks or {}
    xs = [MX]
    for _, w in cols:
        xs.append(xs[-1] + w * mm)
    right = xs[-1]
    cv.setFillGray(0.88)
    cv.rect(MX, top - row_h, right - MX, row_h, stroke=0, fill=1)
    cv.setFillGray(0)
    cv.setFont("Helvetica-Bold", size)
    for x, (lab, _) in zip(xs, cols):
        cv.drawString(x + 1.6 * mm, top - row_h + (row_h - size * 0.7) / 2, lab)
    y = top - row_h
    lines = []
    for i, row in enumerate(rows):
        if i in breaks:
            lines.append((y, breaks[i]))
        y -= row_h
        if shade and i % 2:
            cv.setFillGray(0.965)
            cv.rect(MX, y, right - MX, row_h, stroke=0, fill=1)
            cv.setFillGray(0)
        cv.setFont(body_font, size)
        for x, val in zip(xs, row):
            if val:
                cv.drawString(x + 1.6 * mm, y + (row_h - size * 0.7) / 2, val)
        cv.setStrokeGray(0.7)
        cv.setLineWidth(0.3)
        cv.line(MX, y, right, y)
    cv.setStrokeGray(0.2)
    for ly, lw in lines:
        cv.setLineWidth(lw)
        cv.line(MX, ly, right, ly)
    cv.setStrokeGray(0.35)
    cv.setLineWidth(0.6)
    for k, x in enumerate(xs):
        if k in (0, len(xs) - 1) or k in boxes:
            cv.line(x, top, x, y)
    cv.rect(MX, y, right - MX, top - y, stroke=1, fill=0)
    return y


COVER_COLS = [("Sr", 7), ("TA Roll", 18), ("TA Name", 35), ("Batch/Department", 35),
              ("Room", 13), ("Time Slot", 26), ("Students", 15), ("Evaluation TAs", 33)]
# Columns merged down a run of rows, as on the tab: the key a row must share
# with the row above to stay in the same cell.
MERGED = {4: lambda d: d.room, 5: lambda d: (d.room, d.slot),
          6: lambda d: (d.room, d.slot), 7: lambda d: d.room}


def cover_cells(duties: list[Duty]) -> tuple[list[list[str]], list[tuple[int, int, int, list[str]]]]:
    """Plain per-row values, and merged cells as (col, first row, last row, lines)."""
    plain = [[str(i + 1), d.roll, d.name, d.label] for i, d in enumerate(duties)]
    spans = []
    for col, key in MERGED.items():
        start = 0
        for i in range(1, len(duties) + 1):
            if i == len(duties) or key(duties[i]) != key(duties[start]):
                d = duties[start]
                lines = {4: [d.room], 5: [d.slot], 6: [str(d.slot_students)],
                         7: d.evaluators}[col]
                spans.append((col, start, i - 1, lines))
                start = i
    return plain, spans


def cover_table(cv, top: float, duties: list[Duty], row_h: float, size: float = 8) -> float:
    xs = [MX]
    for _, w in COVER_COLS:
        xs.append(xs[-1] + w * mm)
    right, n = xs[-1], len(duties)
    plain, spans = cover_cells(duties)
    body_top = top - row_h

    cv.setFillGray(0.88)
    cv.rect(MX, body_top, right - MX, row_h, stroke=0, fill=1)
    cv.setFillGray(0)
    cv.setFont("Helvetica-Bold", size)
    for x, (lab, _) in zip(xs, COVER_COLS):
        cv.drawString(x + 1.6 * mm, body_top + (row_h - size * 0.7) / 2, lab)

    cv.setFont("Helvetica", size)
    lead = size * 1.35
    for i, row in enumerate(plain):
        y = body_top - (i + 1) * row_h
        for c, val in enumerate(row):
            # A label too wide for its column breaks before a " + " part.
            width = xs[c + 1] - xs[c] - 3.2 * mm
            lines = [val]
            if cv.stringWidth(val, "Helvetica", size) > width and " + " in val:
                head, _, tail = val.rpartition(" + ")
                lines = [head, "+ " + tail]
            y0 = y + row_h / 2 + (len(lines) - 1) * lead / 2 - size * 0.35
            for k, line in enumerate(lines):
                cv.drawString(xs[c] + 1.6 * mm, y0 - k * lead, line)
    for col, a, b, lines in spans:
        mid = body_top - (a + b + 1) * row_h / 2
        y0 = mid + (len(lines) - 1) * lead / 2 - size * 0.35
        for k, line in enumerate(lines):
            cv.drawString(xs[col] + 1.6 * mm, y0 - k * lead, line)

    # Rules between rows: a boundary inside a merged cell is not drawn in that
    # column. A new room gets a heavy rule, a new slot a medium one.
    for i in range(1, n):
        y = body_top - i * row_h
        prev, cur = duties[i - 1], duties[i]
        weight = (1.4, 0.2) if prev.room != cur.room else \
                 (0.8, 0.3) if prev.slot != cur.slot else (0.3, 0.7)
        cv.setLineWidth(weight[0])
        cv.setStrokeGray(weight[1])
        for c in range(len(COVER_COLS)):
            if c in MERGED and MERGED[c](prev) == MERGED[c](cur):
                continue
            cv.line(xs[c], y, xs[c + 1], y)
    bottom = body_top - n * row_h
    cv.setStrokeGray(0.35)
    cv.setLineWidth(0.6)
    for x in xs:
        cv.line(x, top, x, bottom)
    cv.line(MX, body_top, right, body_top)
    cv.rect(MX, bottom, right - MX, top - bottom, stroke=1, fill=0)
    return bottom


def cover_page(cv, duties: list[Duty], total: int, date: str) -> None:
    top = _header(cv, "TA, room and time slot for showing Quiz 1 answer scripts to students", date)
    y = cover_table(cv, top - 6 * mm, duties, 9.5 * mm)
    cv.setFont("Helvetica", 10)
    y -= 9 * mm
    for line in _wrap(cv, EVAL_NOTE, W - 2 * MX, "Helvetica", 10):
        cv.drawString(MX, y, line)
        y -= 5 * mm
    rooms = len(dict.fromkeys(d.room for d in duties))
    cv.setFont("Helvetica", 8.5)
    cv.drawString(MX, y - 4 * mm, f"Total students: {total} in {rooms} rooms. "
                  "Each row has its own sheet in this pack listing its students.")
    cv.showPage()


def ta_page(cv, sr: int, n: int, d: Duty, groups, date: str) -> None:
    students = students_of(d, groups)
    top = _header(cv, f"Sheet {sr} of {n}   ·   {len(students)} student{'s' if len(students) != 1 else ''}   ·   "
                  "Enter marks by pen before showing the scripts", date)
    cv.setFont("Helvetica-Bold", 11)
    cv.drawString(MX, top - 7 * mm, f"TA: {d.name}  ({d.roll})")
    cv.drawRightString(W - MX, top - 7 * mm, f"Department: {d.label}")
    cv.drawString(MX, top - 13.5 * mm, f"Room: {d.room}   ·   {d.slot}")
    if d.evaluators:
        cv.setFont("Helvetica", 9)
        cv.drawRightString(W - MX, top - 13.5 * mm,
                           "Evaluation TAs: " + ", ".join(d.evaluators))

    footer = 22 * mm
    table_top = top - 17.5 * mm
    rows = [[r[0], r[1], r[2] if len(r[2]) <= 34 else r[2][:33] + ".."] + r[3:]
            for r in sheet_rows(d, groups)]
    row_h = min(7.5 * mm, (table_top - footer) / (len(rows) + 1))
    cols = [("Sr", 10), ("Roll Number", 28), ("Student Name", 62), ("Marks", 22),
            ("Tick", 14), ("Remarks", 46)]
    _table(cv, table_top, cols, rows, row_h, 9 if row_h >= 5.6 * mm else 8.5,
           boxes={3, 4, 5})

    cv.setFont("Helvetica-Oblique", 7.5)
    y = footer - 5 * mm
    for line in _wrap(cv, EVAL_NOTE, W - 2 * MX - 70 * mm, "Helvetica-Oblique", 7.5):
        cv.drawString(MX, y, line)
        y -= 3.6 * mm
    cv.setFont("Helvetica", 9)
    cv.drawRightString(W - MX, footer - 12 * mm, "TA signature: ____________________")
    cv.showPage()


def write_xlsx(path: Path, duties: list[Duty], groups, date: str) -> None:
    wb = Workbook()
    thin = Side(style="thin", color="999999")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    head = PatternFill("solid", fgColor="E0E0E0")

    def rule(s, hdr, ncol):
        for row in s.iter_rows(min_row=hdr, max_row=s.max_row, max_col=ncol):
            if row[0].value is None:
                break
            for c in row:
                c.border = box
                if c.row == hdr:
                    c.font, c.fill = Font(bold=True), head

    ws = wb.active
    ws.title = "TA Assignment"
    ws.append([TITLE, None, None, date])
    ws.append([])
    ws.append(["Sr", "TA Roll", "TA Name", "Batch/Department", "Room", "Time Slot",
               "Students", "Evaluation TAs"])
    plain, spans = cover_cells(duties)
    for row in plain:
        ws.append([int(row[0])] + row[1:] + [None] * 4)
    rule(ws, 3, 8)
    for col, a, b, lines in spans:
        cell = ws.cell(4 + a, col + 1)
        cell.value = int(lines[0]) if col == 6 else " & ".join(lines)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        if b > a:
            ws.merge_cells(start_row=4 + a, start_column=col + 1,
                           end_row=4 + b, end_column=col + 1)
    ws.append([])
    ws.append([EVAL_NOTE])
    for col, w in zip("ABCDEFGH", (5, 13, 26, 24, 9, 20, 10, 24)):
        ws.column_dimensions[col].width = w

    for sr, d in enumerate(duties, 1):
        s = wb.create_sheet(f"{sr:02d} {d.name}"[:31])
        s.append([TITLE, None, None, date])
        s.append([f"TA: {d.name} ({d.roll})", None, f"Department: {d.label}"])
        s.append([f"Room: {d.room}  {d.slot}", None,
                  "Evaluation TAs: " + ", ".join(d.evaluators)])
        s.append(["Sr", "Roll Number", "Student Name", "Marks", "Tick", "Remarks"])
        heads = []
        for row in sheet_rows(d, groups):
            if row[0] == BACKLOG_HEAD:
                s.append([BACKLOG_HEAD])
                heads.append(s.max_row)
            else:
                s.append([int(row[0]), row[1], row[2], None, None, None])
        rule(s, 4, 6)
        for h in heads:
            s.cell(h, 1).font = Font(bold=True, italic=True)
            s.merge_cells(start_row=h, start_column=1, end_row=h, end_column=6)
        for col, w in zip("ABCDEF", (6, 14, 32, 10, 6, 30)):
            s.column_dimensions[col].width = w
    for s in wb.worksheets:
        s["A1"].font = Font(bold=True, size=13)
        s.page_setup.fitToWidth = 1
    wb.save(path)


def _slug(text: str) -> str:
    return re.sub(r"[^\w]+", "_", text).strip("_")


def build(duties: list[Duty], groups: dict[str, pd.DataFrame], total: int, date: str,
          out: Path = OUT) -> dict[str, Path | list[Path]]:
    """Cover + one sheet per duty as one pack PDF, the same sheets split into
    per-TA PDFs, and a tracking workbook. Returns the paths written, so a
    caller (the CLI below, or the dashboard) can report or offer them for
    download without re-deriving the naming scheme."""
    out.mkdir(parents=True, exist_ok=True)
    per_ta = out / "per_ta"
    per_ta.mkdir(exist_ok=True)
    for old in per_ta.glob("*.pdf"):   # numbering follows the tab; drop stale names
        old.unlink()

    pack_path = out / "quiz-1_answer_showing_pack.pdf"
    pack = canvas.Canvas(str(pack_path), pagesize=A4)
    pack.setTitle(TITLE)
    cover_page(pack, duties, total, date)
    per_ta_paths = []
    for sr, d in enumerate(duties, 1):
        ta_page(pack, sr, len(duties), d, groups, date)
        one_path = per_ta / f"{sr:02d}_{d.room}_{_slug(d.slot.split(' ')[0])}_{_slug(d.name)}.pdf"
        one = canvas.Canvas(str(one_path), pagesize=A4)
        one.setTitle(f"{TITLE} · {d.name}")
        ta_page(one, sr, len(duties), d, groups, date)
        one.save()
        per_ta_paths.append(one_path)
    pack.save()

    xlsx_path = out / "quiz-1_answer_showing.xlsx"
    write_xlsx(xlsx_path, duties, groups, date)
    return {"pack": pack_path, "xlsx": xlsx_path, "per_ta": per_ta_paths}


def main() -> None:
    duties, date = read_schedule()
    groups = load_groups()
    check(duties, groups)
    total = sum(len(g) for g in groups.values())
    paths = build(duties, groups, total, date)

    for sr, d in enumerate(duties, 1):
        print(f"{sr:2d}  {d.room:<6} {d.slot:<18} {d.name:<25} {d.label:<24} "
              f"eval: {', '.join(d.evaluators)}")
    print(f"{total} students, {len(duties)} sheets -> {paths['pack'].parent}")


if __name__ == "__main__":
    main()
