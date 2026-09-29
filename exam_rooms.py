"""Exam-hall seating parsed from `LHC Seating Plan.xlsx` - the Institute's
own multi-room exam layout workbook.

This is a deliberately separate model from `config.py`'s classroom rooms:
an exam *venue* can combine several physical rooms into one seating plan -
the workbook's own `LHC 206 207 304 306 307` sheet is one seating plan
spanning five separate rooms - and exam seats are conventionally spaced out
(alternate row, alternate seat) rather than filled solid the way lecture
attendance seating is.

Naming note, confirmed by hand: the source sheets originally used two
different prefixes for the same physical building - everything here has
been normalized to "LHC" for clarity, so `LHC 308`'s own title cell still
reads "Lecture Hall Complex - I" underneath. `LHC110` (the classroom room in
`config.py`) and `LHC 110` here are very likely the same physical room 110,
just two independent spreadsheets for two different purposes (dense daily
lecture seating vs. spaced-out exam seating) - don't assume a mismatch
between their capacities is a bug. "LHC2" (`config.py`'s `LHC2_101`) is a
**different, unrelated building** despite the similar name - never conflate
it with "LHC" here.

Two source layouts are understood:
  * "Block ID - X" header rows (LHC 308's reference sheet) - each block's
    seat-number range for a printed row sits directly under its header, or
    an explicit total in parentheses (`Block ID - A (36)`) for a block
    whose sheet only ever drew a sample row, not its real row count.
  * a flat per-block grid where seats are already labelled "A-15" … "A-1"
    (`LHC 206 207 304 306 307` - one shared layout, independently copied
    into 5 separate rooms of identical capacity, see `_GRID_ROOMS`).
"""
from __future__ import annotations
import re
from dataclasses import dataclass, field

import openpyxl

import config as C
import pandas as pd

SOURCE_XLSX = C.EXAM_HALLS_WORKBOOK
_SKIP_SHEETS = {"sheet5"}          # the example roll-list output, not a room
_GRID_SEAT = re.compile(r"^([A-Za-z])-(\d+)$")


# The three ways a hall can be filled. Which one is right is a decision about
# what you are guarding against, not a technicality:
#   "alternate" - a gap beside AND behind. Nobody has a neighbour in any
#                 direction. Costs three quarters of the hall.
#   "side"      - a gap beside only; every row is used, so a student can sit
#                 behind another. Stops the over-the-shoulder look, which is
#                 the realistic one in a tiered or flat hall, at half the cost.
#   "every"     - every chair. For a hall used only to hold people.
SPACING_MODES = {
    "alternate": "Gap beside and behind (alternate seats, alternate rows)",
    "side": "Gap beside only - every row used, may sit behind",
    "every": "Fill every seat",
}
DEFAULT_SPACING = "alternate"

# Transition seats are held back (see `ExamBlock.reserved`) - but opening a
# whole extra room, with its own invigilator, because one or two students
# don't fit is worse than seating those one or two in the taper. Above this
# many left over, the room really is too small and the honest answer is to
# add another one.
OVERFLOW_LIMIT = 2


def _normalise_mode(mode) -> str:
    """Accepts the old boolean `spaced=` as well as a mode name, so exams
    saved before the third method existed still load."""
    if isinstance(mode, bool):
        return "alternate" if mode else "every"
    mode = str(mode or DEFAULT_SPACING)
    if mode not in SPACING_MODES:
        raise ValueError(f"Unknown spacing mode {mode!r}; expected one of {list(SPACING_MODES)}")
    return mode


@dataclass
class ExamBlock:
    key: str
    rows: list[list[str]] = field(default_factory=list)   # seat numbers, row order, left-to-right
    # Seats that exist but are never allotted. In LHC 110 these are the wedge
    # seats behind the last full row of blocks B and C - the transition pool
    # the classroom side holds back (`config.RESERVE_WEDGE`) for late
    # admissions and mid-term moves. An exam should not seat anyone there
    # either: they are the awkward taper seats at the back corner of a block.
    # Still drawn on the plan, just never filled.
    reserved: set = field(default_factory=set)

    @property
    def dense_seats(self) -> list[str]:
        """Every allottable seat, no spacing."""
        return [s for row in self.rows for s in row if s not in self.reserved]

    @property
    def spaced_seats(self) -> list[str]:
        """Alternate rows, alternate seats within a used row - the spacing
        the Institute's own quiz template (Sheet5) already uses."""
        out = []
        for row in self.rows[::2]:
            out.extend(s for s in row[::2] if s not in self.reserved)
        return out

    @property
    def side_spaced_seats(self) -> list[str]:
        """Alternate seats within *every* row: a free chair on each side, and
        the row behind occupied. Roughly double the capacity of `spaced_seats`
        while still keeping papers out of arm's reach sideways."""
        return [s for row in self.rows for s in row[::2] if s not in self.reserved]

    def seats(self, mode=DEFAULT_SPACING) -> list[str]:
        return {"alternate": self.spaced_seats,
                "side": self.side_spaced_seats,
                "every": self.dense_seats}[_normalise_mode(mode)]

    def overflow_seats(self, mode=DEFAULT_SPACING) -> list[str]:
        """The reserved seats, in row order, spaced the same way as the rest of
        the block - the pool to dip into when one or two students are left
        over. Spacing still applies: an overflow seat beside an allotted one
        would defeat the spacing everyone else got.
        """
        mode = _normalise_mode(mode)
        rows = self.rows[::2] if mode == "alternate" else self.rows
        step = 1 if mode == "every" else 2
        return [s for row in rows for s in row[::step] if s in self.reserved]


@dataclass
class ExamRoom:
    sheet: str
    label: str
    blocks: dict[str, ExamBlock] = field(default_factory=dict)
    supported: bool = True
    note: str = ""

    def capacity(self, mode=DEFAULT_SPACING) -> int:
        return sum(len(b.seats(mode)) for b in self.blocks.values())


_HEADER_RE = re.compile(r"Block ID\s*-\s*([A-Za-z])\s*(?:\((\d+)\))?")


def _parse_block_id_sheet(ws, label: str) -> ExamRoom:
    """`Block ID - X` header row, followed by numbered rows giving each
    block's first/last seat number for that printed row. Handles multiple
    header sections in one sheet (LHC 308 has two: blocks A-D, then E-G).

    A header can instead carry an explicit total - `Block ID - A (36)` -
    for a block whose printed grid only shows a sample row, not every
    physical row (the room has far more rows than were ever drawn). That
    block is fixed at seats 1..N and grid rows below it are ignored."""
    room = ExamRoom(sheet=ws.title, label=label)
    headers: list[tuple[int, str]] = []
    fixed: set[str] = set()
    for row in ws.iter_rows():
        cells = {c.column: c.value for c in row if c.value is not None}
        found = sorted(
            (col, m.group(1).upper(), m.group(2))
            for col, v in cells.items()
            if isinstance(v, str) and (m := _HEADER_RE.match(v.strip()))
        )
        if found:
            headers = [(col, letter) for col, letter, _ in found]
            for col, letter, total in found:
                if total is not None:
                    n = int(total)
                    room.blocks[letter] = ExamBlock(key=letter, rows=[[str(k) for k in range(1, n + 1)]])
                    fixed.add(letter)
                else:
                    room.blocks.setdefault(letter, ExamBlock(key=letter))
            continue
        if not headers:
            continue
        nums = {col: v for col, v in cells.items() if isinstance(v, (int, float))}
        if not nums:
            continue
        for i, (col, letter) in enumerate(headers):
            if letter in fixed:
                continue
            hi_col = headers[i + 1][0] if i + 1 < len(headers) else None
            # Keep left-to-right column order, not numeric order - a spaced
            # allocation skips alternate seats *in the room*, and a block
            # can carry numbers out of numeric sequence (LHC 110 fills gaps
            # left by an earlier partial numbering with numbers appended
            # after the run, not inserted in between).
            span = [v for c, v in sorted(nums.items())
                    if c >= col and (hi_col is None or c < hi_col)]
            if not span:
                continue
            if len(span) <= 2:
                # Sparse "start …→… end" convention (e.g. LHC 308): only the
                # row's first/last seat are written, so fill the gap between.
                lo, hi = int(min(span)), int(max(span))
                seats = [str(n) for n in range(lo, hi + 1)]
            else:
                # Every seat in the row is already written out explicitly.
                seats = [str(int(v)) for v in span]
            room.blocks[letter].rows.append(seats)
    return room


# The sheet draws one room's full layout - 5 blocks (A-E), 15 seats each,
# 75 seats / 40 spaced total - and that same layout is common to 5 separate
# physical rooms with identical seating capacity, not distributed one block
# per room. Each of the 5 gets its own independent copy of all 5 blocks.
_GRID_ROOMS = ["LHC 206", "LHC 207", "LHC 304", "LHC 306", "LHC 307"]

# Rooms the Institute's workbook does not describe. LHC 105 is LHC 206 and
# LHC 207 merged into one hall: ten blocks of fifteen, 150 seats, the same
# block shape those two rooms use, so a row spaced alternately still seats
# eight. Defined here rather than by editing the Institute's own file, which
# this project only ever reads.
EXTRA_ROOMS = {
    "LHC 105": {"rows": 10, "cols": 15},
}

# Rooms with no block letters at all: a seat there is a row and a column, the
# way the hall itself is numbered. LHC 110 keeps blocks - a block there is a
# place an invigilator stands in, with its own grid and its own signature
# sheet - but LHC 105 is one flat hall of ten rows, and calling the third row
# "Block C" only adds a word a student has to translate before sitting down.
# Internally a row is still an `ExamBlock` (keyed "R01" … so that plain sorting
# keeps them in order); every user-facing label goes through `place_label`.
ROW_COL_ROOMS = {"LHC 105"}
_ROW_KEY = re.compile(r"^R(\d+)$")


def is_row_col_room(room) -> bool:
    return str(room).strip() in ROW_COL_ROOMS


def row_number(block) -> int | None:
    """`R03` → 3, and None for a block letter."""
    m = _ROW_KEY.match(str(block).strip())
    return int(m.group(1)) if m else None


def row_key(n: int) -> str:
    return f"R{int(n):02d}"


def group_label(room, block) -> str:
    """What one group of seats is called on paper: `Row 3` in a row/column
    room, `Block C` everywhere else."""
    n = row_number(block)
    if n is not None and is_row_col_room(room):
        return f"Row {n}"
    return f"Block {block}"


def venue_note(room, keys) -> str:
    """How the venue line names the part of a room in use: `Block ID - A & B`
    for a lettered hall, `Rows 1-10` for one numbered by row and column, since
    ten "&"s is not a venue line anybody reads."""
    keys = list(keys)
    if is_row_col_room(room) and all(row_number(k) is not None for k in keys):
        nums = sorted(row_number(k) for k in keys)
        if nums == list(range(nums[0], nums[-1] + 1)):
            return f"Rows {nums[0]}-{nums[-1]}" if len(nums) > 1 else f"Row {nums[0]}"
        return "Rows " + ", ".join(str(n) for n in nums)
    return "Block ID - " + " & ".join(str(k) for k in keys)


def place_label(room, block, seat=None) -> str:
    """A whole address: `LHC 105, Row 3, Seat 7` or `LHC 110, Block C, Seat 12`.
    Seat is the column number in a row/column room."""
    parts = [str(room)] if room else []
    if str(block).strip():
        parts.append(group_label(room, block))
    if seat is not None and str(seat).strip():
        parts.append(("Column " if is_row_col_room(room) else "Seat ") + str(seat))
    return ", ".join(parts)

# Rooms the workbook still lists but the course no longer uses. They stay
# visible with the reason attached rather than disappearing, so that anyone
# looking for LHC 206 finds out where it went instead of assuming the parser
# broke.
SUPERSEDED = {
    "LHC 206": "Merged into LHC 105 (150 seats). Pick LHC 105 instead.",
    "LHC 207": "Merged into LHC 105 (150 seats). Pick LHC 105 instead.",
}


def _room_sort_key(label: str):
    """LHC 105 before LHC 110 before LHC 308: read as room numbers, not as
    text, and keep anything unnumbered at the end."""
    digits = "".join(ch for ch in str(label) if ch.isdigit())
    return (0, int(digits)) if digits else (1, 0)


def _extra_rooms() -> dict[str, "ExamRoom"]:
    """A row/column room is one `ExamBlock` per physical row, each holding that
    row's columns. Keeping a row as a block is what preserves the spacing
    arithmetic: alternate spacing skips every second column of every row, which
    in a ten-row hall of fifteen is the eighty places the two rooms this one
    replaced used to give between them."""
    rooms = {}
    for label, spec in EXTRA_ROOMS.items():
        blocks = {}
        if "rows" in spec:
            seats = [str(i) for i in range(1, int(spec["cols"]) + 1)]
            for n in range(1, int(spec["rows"]) + 1):
                blocks[row_key(n)] = ExamBlock(key=row_key(n), rows=[list(seats)])
        else:
            for key in spec["blocks"]:
                seats = [str(i) for i in range(1, int(spec["per_block"]) + 1)]
                blocks[key] = ExamBlock(key=key, rows=[seats])
        rooms[label] = ExamRoom(sheet="(defined in exam_rooms.EXTRA_ROOMS)",
                                label=label, blocks=blocks)
    return rooms


def _parse_grid_sheet(ws) -> dict[str, ExamRoom]:
    """One row per block, seats already spelled out as `A-15` … `A-1`. The
    full block set (A-E) is the layout shared by all 5 physical rooms in
    `_GRID_ROOMS` - each gets an independent copy of it, not one block each."""
    blocks: dict[str, ExamBlock] = {}
    for row in ws.iter_rows():
        vals = [c.value for c in row]
        letters = {v.strip() for v in vals if isinstance(v, str) and re.fullmatch(r"[A-Za-z]", v.strip())}
        seats = [m for v in vals if isinstance(v, str) and (m := _GRID_SEAT.match(v.strip()))]
        if len(letters) == 1 and seats:
            letter = next(iter(letters))
            nums = sorted({s.group(2) for s in seats}, key=int)
            blocks[letter] = ExamBlock(key=letter, rows=[nums])
    if not blocks:
        return {}
    rooms: dict[str, ExamRoom] = {}
    for label in _GRID_ROOMS:
        room = ExamRoom(sheet=ws.title, label=label)
        room.blocks = {k: ExamBlock(key=k, rows=[list(r) for r in b.rows]) for k, b in blocks.items()}
        rooms[label] = room
    return rooms


def _known_geometry(room_label: str) -> dict[str, "C.Block"] | None:
    """The same hall, if this project already knows its real shape.

    `LHC 110` in the exam workbook is the hall `config.ROOMS["LHC110"]`
    describes seat-for-seat (A 6×6=36, B 42+8 wedge=50, C 42+12=54, D 36,
    E 14×10=140, F 13×10=130, G 140) - but the exam sheet gives only each
    block's *total*, so without this the block is one flat row of 140.

    That is not just an ugly drawing. `spaced_seats` skips alternate rows and
    alternate seats *within* a row; over a single 140-long row there are no
    rows to skip, so "alternate seating" degenerates into every second chair
    in one continuous line - which in a ten-wide hall seats people directly
    in front of and behind each other.
    """
    key = room_label.replace(" ", "")
    room = C.ROOMS.get(key)
    return {b.key: b for b in room} if room else None


def _shape_from_geometry(seats: list[str], block: "C.Block") -> tuple[list[list[str]], set] | None:
    """Chop a flat seat list into the block's real rows: full rows first, then
    the wedge/taper rows behind them. Returns None if the totals disagree -
    a mismatch means the two sources describe different halls, and guessing
    would be worse than falling back to a plain wrap."""
    if len(seats) != block.total:
        return None
    rows, i = [], 0
    for _ in range(block.rows):
        rows.append(seats[i:i + block.cols])
        i += block.cols
    reserved: set = set()
    for extra in block.wedge:
        wedge_row = seats[i:i + extra]
        rows.append(wedge_row)
        reserved.update(wedge_row)          # drawn, never allotted
        i += extra
    return rows, reserved


def _apply_known_geometry(rooms: dict[str, ExamRoom]) -> None:
    """Reshape flat blocks in place wherever `config.ROOMS` knows the hall."""
    for label, room in rooms.items():
        geom = _known_geometry(label)
        if not geom:
            continue
        for key, block in room.blocks.items():
            if len(block.rows) > 1:
                continue                      # the sheet already gave real rows
            # `dense_seats` already filters reserved seats, so reshape from the
            # raw rows - otherwise a second call would lose the wedge entirely.
            flat = [s for row in block.rows for s in row]
            shaped = _shape_from_geometry(flat, geom[key]) if key in geom else None
            if shaped:
                block.rows, block.reserved = shaped


def load_exam_rooms(path: str = SOURCE_XLSX) -> dict[str, ExamRoom]:
    """One entry per room/venue found in the workbook, keyed by its label."""
    wb = openpyxl.load_workbook(path, data_only=True)
    rooms: dict[str, ExamRoom] = {}
    for name in wb.sheetnames:
        if name.lower().startswith("copy of") or name.lower() in _SKIP_SHEETS:
            continue
        ws = wb[name]
        a1 = ws.cell(1, 1).value
        title = str(a1).strip() if a1 else name

        has_block_header = any(
            isinstance(c.value, str) and c.value.strip().startswith("Block ID -")
            for r in ws.iter_rows() for c in r
        )
        if has_block_header:
            m = re.search(r"LHC\s*\d+", title, re.I)
            label = m.group(0) if m else title
            room = _parse_block_id_sheet(ws, label)
            if room.blocks:
                rooms[room.label] = room
            continue

        has_grid = any(
            isinstance(c.value, str) and _GRID_SEAT.match(c.value.strip())
            for r in ws.iter_rows() for c in r
        )
        if has_grid:
            rooms.update(_parse_grid_sheet(ws))
            continue

        rooms[name] = ExamRoom(
            sheet=name, label=name, blocks={}, supported=False,
            note="No filled-in seat numbers found in this sheet - looks like an "
                 "unfinished layout template, so it's left out of allocation.",
        )
    rooms.update(_extra_rooms())
    _apply_known_geometry(rooms)
    for label, why in SUPERSEDED.items():
        if label in rooms:
            rooms[label].supported = False
            rooms[label].note = why
    return {k: rooms[k] for k in sorted(rooms, key=_room_sort_key)}


def build_exam_roll_list(alloc, out_xlsx: str, *, institute=C.INSTITUTE,
                         exam_title: str, course: str, datetime_label: str,
                         venue_label: str, ta: str = "", room_map: dict[str, "ExamRoom"] | None = None) -> str:
    """Writes the Institute's own exam roll-list format: header block, then
    S.No / Roll No. / Name / Room / Block ID / Seat No. / Attendance.

    `alloc` needs a `Room` column alongside `Block`/`Seat`/`Roll`/`Name` -
    students can span more than one physical room. If `room_map` (the
    `ExamRoom` objects the seats were drawn from) is given, a second
    "Layout" sheet is added: one line per physical seating row, arranged
    left-to-right the way the room actually reads, with each seat showing
    its roll number (or "-" for a seat that exists but wasn't allotted -
    e.g. skipped for spacing, or left over capacity)."""
    from openpyxl import Workbook
    from openpyxl.styles import Font

    wb = Workbook()
    ws = wb.active
    ws.title = "Seating"
    bold = Font(bold=True)
    for line in (institute, exam_title, course, datetime_label, f"Exam Venue : {venue_label}", f"TA {ta}"):
        ws.append([line])
        ws.cell(ws.max_row, 1).font = bold if ws.max_row == 1 else Font()
    ws.append([])
    headers = ["S. No.", "Roll No.", "Name", "Room", "Block ID", "Seat No.", "Attendance"]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        ws.cell(ws.max_row, c).font = bold
    for i, row in enumerate(alloc.itertuples(index=False), start=1):
        ws.append([i, row.Roll, row.Name, row.Room, row.Block, row.Seat, None])
    for col, width in zip("ABCDEFG", (7, 14, 26, 22, 10, 10, 12)):
        ws.column_dimensions[col].width = width

    if room_map:
        _add_layout_sheet(wb, alloc, room_map)

    wb.save(out_xlsx)
    return out_xlsx


_MATRIX_WIDTH = 15  # wrap any physical row wider than this into a proper grid


def _wrap_rows(phys_rows: list[list[str]], width: int = _MATRIX_WIDTH) -> list[list[str]]:
    # `width` is a maximum, not a target: a row already narrower than it keeps
    # its real shape, so passing a smaller width only re-wraps the flat rows.
    """A block's rows are usually already narrow (LHC 308 tops out at 11
    seats/row), but a block known only by its *total* (LHC 110's back
    section, or LHC's flat 15-seat row) is stored as one long row - wrap
    those into a proper multi-row grid instead of one very long line."""
    out = []
    for row in phys_rows:
        if len(row) > width:
            out.extend(row[i:i + width] for i in range(0, len(row), width))
        else:
            out.append(row)
    return out


def _add_layout_sheet(wb, alloc, room_map: dict[str, "ExamRoom"]) -> None:
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    ws = wb.create_sheet("Layout")
    bold = Font(bold=True)
    title_fill = PatternFill("solid", fgColor="DCE6F1")
    hdr_fill = PatternFill("solid", fgColor="F2F2F2")
    thin = Side(style="thin", color="999999")
    box = Border(thin, thin, thin, thin)
    centre = Alignment(horizontal="center", vertical="center")
    roll_at = {(r.Room, r.Block, str(r.Seat)): r.Roll for r in alloc.itertuples(index=False)}

    out_row = 1
    max_col_used = 1
    for room_label in pd.unique(alloc.Room):
        ws.cell(out_row, 1, room_label).font = Font(bold=True, size=13)
        out_row += 1
        room = room_map.get(room_label)
        block_order = list(pd.unique(alloc.loc[alloc.Room == room_label, "Block"]))
        for bk in block_order:
            phys_rows = room.blocks[bk].rows if room and bk in room.blocks else None
            if not phys_rows:
                phys_rows = [list(alloc.loc[(alloc.Room == room_label) & (alloc.Block == bk), "Seat"].astype(str))]
            grid = _wrap_rows(phys_rows)
            width = max(len(r) for r in grid)
            max_col_used = max(max_col_used, width + 1)

            ws.cell(out_row, 1, f"Block {bk}").font = bold
            ws.cell(out_row, 1).fill = title_fill
            for j in range(2, width + 2):
                c = ws.cell(out_row, j)
                c.fill = title_fill
                c.border = box
            out_row += 1

            for j in range(2, width + 2):
                c = ws.cell(out_row, j, j - 1)
                c.font = bold
                c.fill = hdr_fill
                c.border = box
                c.alignment = centre
            out_row += 1

            for r_i, phys_row in enumerate(grid, start=1):
                c = ws.cell(out_row, 1, f"Row {r_i}")
                c.font = bold
                c.fill = hdr_fill
                c.border = box
                for j, seat in enumerate(phys_row, start=2):
                    cell = ws.cell(out_row, j, roll_at.get((room_label, bk, seat), "-"))
                    cell.border = box
                    cell.alignment = centre
                out_row += 1
            out_row += 1  # blank line between blocks
        out_row += 1  # blank line between rooms

    from openpyxl.utils import get_column_letter
    ws.column_dimensions["A"].width = 16
    for col in range(2, max_col_used + 1):
        ws.column_dimensions[get_column_letter(col)].width = 12
