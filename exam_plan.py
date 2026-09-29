#!/usr/bin/env python3
"""Exam seating as something that persists, prints and can be looked up.

The classroom side of this project has an allocation on disk, printable
sheets, a published lookup page and a mail-out. The exam side had a
spreadsheet and nothing else: the allocation lived in one Streamlit session
and vanished with it, and a student's only route to their seat was a roll list
taped to a wall.

This module gives an exam the same footing:

    out/exams/<slug>/allocation.csv   who sits where - the source of truth
    out/exams/<slug>/exam.json        title, date, venue, spacing, rooms
    out/exams/<slug>/seat_map.html    invigilator/notice-board plan, searchable
    out/exams/<slug>/find-my-seat.html  student lookup, roll numbers hashed

Seats are drawn in the room's real geometry - `exam_rooms.ExamBlock.rows` is
the physical row structure parsed from the Institute's own sheet - so a
seat on screen is in the same place as the seat in the hall.
"""
from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

import config as C
import finder

EXAM_ROOT = Path("out") / "exams"

# The term's scheduled exams, as fixed by the department. Both are sat by both
# batches together - English and Hindi are one cohort for an exam, which is
# why the exam side has always been a separate model from the classroom side
# rather than something per-medium.
SCHEDULE = [
    {"key": "quiz1", "title": "Quiz 1",
     "date": "31 August 2026", "day": "Monday",
     "time": "08:00 PM, 30–45 minutes", "batches": "English + Hindi"},
    {"key": "minor", "title": "Minor Exam",
     "date": "16 September 2026", "day": "Wednesday",
     "time": "04:30 PM to 06:30 PM", "batches": "English + Hindi"},
    {"key": "major", "title": "Major Exam",
     "date": "21 November 2026", "day": "Saturday",
     "time": "06:00 PM to 09:00 PM", "batches": "English + Hindi"},
]


def scheduled(key: str) -> dict | None:
    return next((e for e in SCHEDULE if e["key"] == key), None)


def when_label(exam: dict) -> str:
    """`16 September 2026, Wednesday (04:30 PM to 06:30 PM)` - the form the
    Institute's own roll-list header uses."""
    return f"{exam['date']}, {exam['day']} ({exam['time']})"
META_FILE = "exam.json"
ALLOC_FILE = "allocation.csv"


def slugify(text: str) -> str:
    s = re.sub(r"[^a-zA-Z0-9]+", "-", str(text).strip().lower()).strip("-")
    return s or "exam"


def exam_roster(roll_list: str, workbook: str | None = None) -> pd.DataFrame:
    """The exam roster in the Institute's own roll-list order, with anyone the
    export has missed put **first**.

    The roll list is the order the department reads students in, so an exam
    that follows it is checkable against their own paperwork. But the export
    is a snapshot - B22PH903 is on the register and in class, and simply isn't
    in it. Dropping him would leave a student with no seat; appending him
    silently at the end would bury a late admission in the last row of the
    last block. He goes at the front, where whoever prints the sheets sees him.
    """
    import seating

    roll = seating.load_roster(roll_list)
    roll["Roll"] = roll.Roll.astype(str).str.strip()
    order = list(roll.Roll)

    if workbook is None:
        workbook = seating.ATTENDANCE_WORKBOOK
    if not Path(workbook).exists():
        return roll.reset_index(drop=True)

    register = seating.load_attendance_roster(workbook)
    register["Roll"] = register.Roll.astype(str).str.strip()
    missing = register[~register.Roll.isin(order)]
    if missing.empty:
        return roll.reset_index(drop=True)

    cols = [c for c in roll.columns if c in missing.columns] or ["Roll", "Name"]
    head = missing[[c for c in ("Roll", "Name", "Medium") if c in missing.columns]].copy()
    for c in roll.columns:
        if c not in head.columns:
            head[c] = ""
    return pd.concat([head[roll.columns], roll], ignore_index=True)


# ────────────────────────────── persistence ──────────────────────────────
def save_exam(alloc: pd.DataFrame, *, title: str, when: str = "", venue: str = "",
              course: str = C.COURSE, spaced="alternate",
              slug: str = "", root: Path = EXAM_ROOT) -> Path:
    """Write the allocation and its metadata. Returns the exam's directory.

    Overwrites an existing exam of the same slug on purpose - re-allocating
    the same quiz should replace it, not accumulate near-duplicates nobody can
    tell apart. The dashboard snapshots first, so the previous one is
    recoverable.
    """
    slug = slug or slugify(title)
    d = Path(root) / slug
    d.mkdir(parents=True, exist_ok=True)
    alloc.to_csv(d / ALLOC_FILE, index=False)
    (d / META_FILE).write_text(json.dumps({
        "slug": slug, "title": title, "when": when, "venue": venue, "course": course,
        # `spaced` is a mode name ("alternate" / "side" / "every"); older
        # exams saved a bool, which `exam_rooms._normalise_mode` still accepts.
        "spaced": spaced, "students": int(len(alloc)),
        "rooms": sorted(set(alloc.Room)) if "Room" in alloc else [],
        "saved": datetime.now().isoformat(timespec="seconds"),
    }, indent=1), encoding="utf-8")
    return d


def list_exams(root: Path = EXAM_ROOT) -> list[dict]:
    """Newest first, each with its metadata and directory."""
    out = []
    for d in sorted(Path(root).glob("*")) if Path(root).exists() else []:
        meta_path, alloc_path = d / META_FILE, d / ALLOC_FILE
        if not alloc_path.exists():
            continue
        meta = {}
        if meta_path.exists():
            try:
                meta = json.loads(meta_path.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                meta = {}
        meta.setdefault("slug", d.name)
        meta.setdefault("title", d.name)
        meta["dir"] = d
        out.append(meta)
    return sorted(out, key=lambda m: m.get("saved", ""), reverse=True)


def load_exam(slug: str, root: Path = EXAM_ROOT) -> tuple[pd.DataFrame, dict]:
    d = Path(root) / slug
    alloc = pd.read_csv(d / ALLOC_FILE, dtype={"Roll": str})
    meta = json.loads((d / META_FILE).read_text(encoding="utf-8")) if (d / META_FILE).exists() else {}
    meta["dir"] = d
    return alloc, meta


# ─────────────────────────── geometry for drawing ────────────────────────
def room_layout(alloc: pd.DataFrame, room_map: dict) -> dict:
    """{room: {block: [[{seat, roll, used}, …rows…]]}} - the real row/seat
    geometry with the allocation laid over it.

    A seat the exam didn't use is still drawn, greyed: the empty chairs are
    what make a spaced layout legible from the front of the hall, and an
    invigilator counting heads needs to see them.
    """
    seat_of = {}
    for r in alloc.itertuples():
        seat_of[(str(getattr(r, "Room", "")), str(r.Block), str(r.Seat))] = str(r.Roll)

    import exam_rooms

    out: dict[str, dict] = {}
    for room_label in (dict.fromkeys(alloc.Room) if "Room" in alloc else []):
        room = room_map.get(room_label)
        if room is None:
            continue
        blocks = {}
        for key, block in room.blocks.items():
            # `exam_rooms` reshapes a block to its real geometry when this
            # project knows the hall; this wrap is the fallback for one it
            # doesn't, so no block is ever drawn as a single endless line.
            reserved = getattr(block, "reserved", set())
            grid = []
            for row in exam_rooms._wrap_rows(block.rows):
                grid.append([{"seat": s, "roll": seat_of.get((room_label, key, s), ""),
                              "used": (room_label, key, s) in seat_of,
                              # drawn but never allottable - the transition
                              # pool, marked so nobody wonders why it's empty
                              "held": s in reserved} for s in row])
            if any(cell["used"] for row in grid for cell in row):
                blocks[key] = grid
        if blocks:
            out[room_label] = blocks
    return out


# ──────────────────────────────── HTML out ───────────────────────────────
_SEAT_MAP_TPL = "exam_seat_map_template.html"
_FINDER_TPL = "exam_finder_template.html"


def _render(template: str, payload: dict, out_html: str) -> str:
    tpl = (Path(__file__).parent / template).read_text(encoding="utf-8")
    Path(out_html).write_text(
        tpl.replace("__DATA__", json.dumps(payload, separators=(",", ":"), ensure_ascii=False)),
        encoding="utf-8")
    return out_html


def _labels(alloc) -> dict:
    """{room: {block: "Row 3" / "Block C"}} plus the word each room uses for a
    seat. LHC 110 is read as blocks; LHC 105 is read as rows and columns, and
    the pages have to say what the hall itself says."""
    import exam_rooms
    out = {}
    for room in dict.fromkeys(alloc.Room) if "Room" in alloc else []:
        sub = alloc[alloc.Room == room]
        out[str(room)] = {
            "rowcol": bool(exam_rooms.is_row_col_room(room)),
            "unit": "column" if exam_rooms.is_row_col_room(room) else "seat",
            "group": {str(b): exam_rooms.group_label(room, b)
                      for b in dict.fromkeys(sub.Block)},
        }
    return out


def build_seat_map(alloc: pd.DataFrame, room_map: dict, out_html: str, *, meta: dict) -> str:
    """The invigilator / notice-board plan: every room drawn to scale, roll
    numbers in position, search box to spotlight one. Not for publishing - it
    lists roll numbers in the clear."""
    payload = {
        "title": meta.get("title", "Exam"), "course": meta.get("course", ""),
        "when": meta.get("when", ""), "venue": meta.get("venue", ""),
        "spaced": meta.get("spaced", "alternate"), "n": int(len(alloc)),
        "instructors": C.instructor_names(),
        "rooms": room_layout(alloc, room_map),
        "labels": _labels(alloc),
        "summary": [{"room": room, "block": block,
                     "label": __import__("exam_rooms").group_label(room, block),
                     "n": int(((alloc.Room == room) & (alloc.Block == block)).sum()),
                     "first": str(sub.Roll.iloc[0]), "last": str(sub.Roll.iloc[-1])}
                    for room in dict.fromkeys(alloc.Room)
                    for block in dict.fromkeys(alloc[alloc.Room == room].Block)
                    for sub in [alloc[(alloc.Room == room) & (alloc.Block == block)]
                                .sort_values("Seat")]],
    }
    return _render(_SEAT_MAP_TPL, payload, out_html)


def build_finder(alloc: pd.DataFrame, out_html: str, *, meta: dict,
                 room_map: dict | None = None) -> str:
    """The student-facing page: type a roll number, get room + block + seat,
    with that seat marked inside a drawing of its block.

    Roll numbers are hashed with the same one-way function the classroom
    lookup page uses, and the block drawings carry **seat labels only, never
    roll numbers** - so the page can sit on a public URL or behind a QR code
    by the hall door without becoming a way to read off who sits where.
    """
    lookup = {finder.roll_key(r.Roll): f"{getattr(r, 'Room', '')}|{r.Block}|{r.Seat}"
              for r in alloc.itertuples()}

    shapes: dict[str, dict] = {}
    if room_map:
        for room, blocks in room_layout(alloc, room_map).items():
            shapes[room] = {b: [[{"s": c["seat"], "u": c["used"], "h": c["held"]} for c in row]
                                for row in grid]
                            for b, grid in blocks.items()}

    payload = {
        "title": meta.get("title", "Exam"), "course": meta.get("course", ""),
        "when": meta.get("when", ""), "venue": meta.get("venue", ""),
        "lookup": lookup, "shapes": shapes, "labels": _labels(alloc),
        "instructors": C.instructor_names(),
        "rooms": [{"room": room,
                   "blocks": [{"k": b, "n": int(((alloc.Room == room) & (alloc.Block == b)).sum())}
                              for b in dict.fromkeys(alloc[alloc.Room == room].Block)]}
                  for room in dict.fromkeys(alloc.Room)],
    }
    return _render(_FINDER_TPL, payload, out_html)


def build_all(alloc: pd.DataFrame, room_map: dict, *, meta: dict,
              root: Path = EXAM_ROOT) -> dict[str, str]:
    """Save the exam and write both pages beside it."""
    d = save_exam(alloc, title=meta.get("title", "Exam"), when=meta.get("when", ""),
                  venue=meta.get("venue", ""), course=meta.get("course", C.COURSE),
                  spaced=meta.get("spaced", True), slug=meta.get("slug", ""), root=root)
    return {
        "dir": str(d),
        "allocation": str(d / ALLOC_FILE),
        "seat_map": build_seat_map(alloc, room_map, str(d / "seat_map.html"), meta=meta),
        "finder": build_finder(alloc, str(d / "find-my-seat.html"), meta=meta,
                               room_map=room_map),
    }


# ─────────────────────────── seating strategy ────────────────────────────
def branch_of(roll: str) -> str:
    """`B26CS1542` → `CS`. Same convention as `seating.branch_code`."""
    m = re.match(r"^[A-Za-z]\d{2}([A-Za-z]{2})", str(roll).strip())
    return m.group(1).upper() if m else "??"


def interleave_by_branch(rolls: list[str]) -> list[str]:
    """Reorder so neighbours are rarely from the same branch.

    A roster sorted by roll number seats CS next to CS for forty seats; the
    people most likely to have prepared together, and to have the same paper
    layout in their heads, end up side by side. Dealing round-robin from
    per-branch piles breaks that up without any randomness - the same roster
    always produces the same plan, which matters when a query comes in a week
    later.
    """
    piles: dict[str, list[str]] = {}
    for r in rolls:
        piles.setdefault(branch_of(r), []).append(r)
    order = sorted(piles, key=lambda b: (-len(piles[b]), b))
    out: list[str] = []
    while any(piles[b] for b in order):
        for b in order:
            if piles[b]:
                out.append(piles[b].pop(0))
    return out
