#!/usr/bin/env python3
"""Two print-and-stick PDFs that don't belong to any other module.

`hall_poster` - an A4 sheet for the hall door: the block map, the lookup URL
in large type, and a QR code to the same. A student at the door with a phone
should not have to type a URL, and a student without one should not be stuck.

`invigilator_pack` - for an exam: one page per room holding the seat grid with
roll numbers in position, followed by a signature list per block. This is what
the exam side has been missing; the classroom side has had signature sheets
from the start.
"""
from __future__ import annotations

import io
import re
from pathlib import Path

from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

import config as C

PAGE_W, PAGE_H = 210.0, 297.0
# A3 landscape for the notice board. A seat map that a student reads from
# arm's length in a corridor is a different object from the A4 grid an
# invigilator holds: it shows the whole room at once, and the roll numbers
# have to be legible with people standing in front of it.
A3_W, A3_H = 420.0, 297.0
# Only the big halls need A3. A five-block room with forty students fits an A4
# sheet at the same cell size, and a board full of half-empty A3 is harder to
# read, not easier.
A3_ROOMS = {"LHC 110", "LHC 308", "LHC 105"}
LEGAL_W, LEGAL_H = 355.6, 215.9  # 14 x 8.5 in (Landscape Legal)


# Which rooms are laid out block by block. LHC 110 and LHC 308 are big enough
# that a block is a place you stand in - each has its own invigilator, and its
# own grid and signature sheet. The smaller rooms are one line of sight with
# two TAs for the room, and eight students per block there would mean five
# grids and five signature sheets for forty people; those print room-wise, the
# blocks marked inside.
BLOCK_WISE_ROOMS = {"LHC 110", "LHC 308"}

# A grid wider than this many columns prints landscape in the pack: past it,
# portrait A4 shrinks a desk below the width a roll number needs.
LANDSCAPE_COLS = 12

# How many names fit down one portrait signature sheet. Past this the list
# prints two columns to a landscape sheet instead (see `_signature_two_up`):
# the alternative is an invigilator holding three pages for one block.
TWO_UP_OVER = 26

# LHC 110's own column numbering, read straight from `LHC110.xlsx`'s header
# rows (row 7 for the front section D-C-B-A, row 17 for the rear G-F-E) and
# cross-checked cell-for-cell against the uploaded seating-plan HTML. Printed
# beside the seat grid so a student can find "Block C, Column 9" the way the
# room itself is signed, rather than by counting boxes across the page.
# Block E's own last entry, "C 11" instead of the "C 32" the sequence would
# give, is not a transcription slip here - it is what that header cell holds
# in the workbook itself, kept verbatim rather than silently corrected.
LHC110_COLUMNS = {
    "D": [1, 2, 3, 4, 5, 6],
    "C": [7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17],
    "B": [18, 19, 20, 21, 22, 23, 24, 25, 26, 27],
    "A": [28, 29, 30, 31, 32, 33],
    "G": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
    "F": [11, 12, 13, 14, 15, 16, 17, 18, 19, 20],
    "E": [22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 11],
}


def _lhc110_columns(room, block_key):
    """This block's column numbers, or None everywhere but LHC 110 - so
    every other hall keeps drawing exactly as it did before."""
    if str(room).strip() != "LHC 110":
        return None
    return LHC110_COLUMNS.get(str(block_key))


def _is_block_wise(room: str) -> bool:
    import exam_rooms
    return (str(room).strip() in BLOCK_WISE_ROOMS
            and not exam_rooms.is_row_col_room(room))


def _is_row_col(room: str) -> bool:
    """A hall printed as rows and columns rather than as lettered blocks."""
    import exam_rooms
    return exam_rooms.is_row_col_room(room)


def _row_no(block) -> str:
    import exam_rooms
    n = exam_rooms.row_number(block)
    return str(n) if n is not None else str(block)


def _group_label(room, block) -> str:
    import exam_rooms
    return exam_rooms.group_label(room, block)


# A spaced hall is mostly empty chairs: with alternate seating every second
# column has nobody in it. Drawn full size those blanks eat the width the roll
# numbers need, so an unused column is printed at this fraction of a used one,
# and an unused *row* at the matching fraction of the row height. The gaps stay
# visible - an invigilator counting heads needs to see which chairs should be
# empty - they just stop being the loudest thing on the sheet.
NARROW_COL = 0.30
EMPTY_ROW = 0.45
# A desk box only has to hold one short line of type, so height past this
# fraction of its width is white space that pushes the roll number down.
FLAT_CELL = 0.5


def _seated_cols(rows, key, seat_roll) -> list[bool]:
    """Which columns of a block anybody actually sits in."""
    cols = max((len(r) for r in rows), default=0)
    return [any(seat_roll.get((str(key), str(r[i])), "") for r in rows if i < len(r))
            for i in range(cols)]


def _col_weight(seated) -> float:
    """How wide a block is, in units of one used column."""
    return sum(1.0 if u else NARROW_COL for u in seated) or 1.0


def _col_widths(seated, wide: float) -> list[float]:
    return [wide if u else wide * NARROW_COL for u in seated]


def _row_is_empty(row, key, seat_roll) -> bool:
    return not any(seat_roll.get((str(key), str(s)), "") for s in row)


def _draw_seat_row(cv, row, *, key, x0, y, widths, cell_h, seat_roll, reserved,
                   roll_font, seat_font, pad=0.9) -> None:
    """One physical row of desks. A filled desk is a full box with the roll
    number in it; an empty one is a small centred box with just its seat
    number, so the gap reads as a gap rather than as a seat somebody forgot."""
    for i, seat in enumerate(row):
        if i >= len(widths):
            break
        if seat is None or str(seat).strip() in ("", "None"):
            continue
        x = x0 + sum(widths[:i])
        w = widths[i] - pad
        roll = seat_roll.get((str(key), str(seat)), "")
        held = str(seat) in reserved
        h = cell_h if roll else cell_h * EMPTY_ROW
        y_box = y if roll else y + (cell_h - h) / 2
        cv.setLineWidth(0.45 if roll else 0.35)
        if roll:
            cv.setFillColorRGB(0.90, 0.95, 0.94)
            cv.setStrokeColorRGB(0.06, 0.46, 0.43)
        elif held:
            cv.setFillColorRGB(0.94, 0.94, 0.92)
            cv.setStrokeColorRGB(*RULE)
        else:
            cv.setFillColorRGB(1, 1, 1)
            cv.setStrokeColorRGB(*RULE)
        cv.rect(x * mm, y_box * mm, w * mm, h * mm, stroke=1, fill=1)
        if roll:
            # Seat number now prints bold and close to the roll number's own
            # size, top-left, so it reads at arm's length rather than needing
            # a second look - the plan is supposed to save people that step.
            cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica-Bold", seat_font)
            cv.drawString((x + 1.0) * mm, (y + cell_h - seat_font / 2.1 - 0.9) * mm,
                          str(seat))
            cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", roll_font)
            cv.drawCentredString((x + w / 2) * mm,
                                 (y + cell_h / 2 - roll_font / 3.4) * mm, roll)
        else:
            # Fit against this box's own (now much narrower) width rather
            # than scaling off the filled cell's seat_font - a bold, blown-up
            # corner label next door must not leave the vacant number too big
            # for its now-smaller box.
            label = "held" if held else str(seat)
            small = min(seat_font * 0.55, cell_h * 3.4)
            while small > 3.0 and cv.stringWidth(label, "Helvetica", small) / mm > (w - 0.8):
                small -= 0.2
            cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", small)
            cv.drawCentredString((x + w / 2) * mm,
                                 (y_box + h / 2 - small / 3.2) * mm,
                                 label)


def _fit_font(cv, text_width_mm: float, sample: str, cell_h: float,
              cap: float = 13.0) -> tuple[float, float]:
    """Type that grows with the box but is measured against the widest roll
    number, not estimated - an estimate is what let a roll run past its cell."""
    f = min(cap, cell_h * 1.25)
    while f > 4.2 and cv.stringWidth(sample, "Helvetica-Bold", f) / mm > text_width_mm:
        f -= 0.2
    # Seat number: close to the roll number's own size and drawn bold (see
    # _draw_seat_row) - a corner label you have to lean in to read is not
    # doing its job, and the vacant columns were narrowed to make room for
    # exactly this.
    return f, max(6.0, min(12.5, f * 0.86))


def _rowcol_grid(cv, sub, room, blocks, *, x0, y_top, avail_w, avail_h,
                 seat_roll, row_gap=1.6, min_cell_h=0.0) -> float:
    """One room drawn the way it is numbered: column numbers across the top,
    row numbers down the left, every desk in its real place.

    A student reading this only has to count along a row, which is what the
    hall itself lets them do.

    An empty column is drawn narrow and short. It still has to be there - the
    gaps are what make a spaced hall legible, and an invigilator counting heads
    needs to see which chairs should be empty - but a blank box the size of a
    filled one spends half the paper saying nothing, and that width is exactly
    what the roll numbers want. Boxes are also kept flatter than they are wide:
    a roll number is one short line, so the height above and below it is dead
    space that pushes the type down.

    Returns the y of the bottom of the last row.
    """
    keys = [k for k in sorted(blocks) if (sub.Block.astype(str) == str(k)).any()]
    if not keys:
        return y_top
    grids = {k: (blocks[k].rows if blocks[k] is not None
                 else [sorted(sub[sub.Block.astype(str) == str(k)].Seat,
                              key=lambda v: int(re.sub(r"\D", "", str(v)) or 0))])
             for k in keys}
    # a row of a row/column room is one physical row; if a block ever arrives
    # with several, they simply stack under the same row label
    lines = [(k, row) for k in keys for row in grids[k]]
    cols = max(len(r) for _, r in lines)

    seated = [any(seat_roll.get((str(k), str(row[i])), "")
                  for k, row in lines if i < len(row))
              for i in range(cols)]

    gutter, head_h = 12.0, 5.0
    wide_w = (avail_w - gutter) / _col_weight(seated)
    col_w = _col_widths(seated, wide_w)
    x_at = [x0 + gutter + sum(col_w[:i]) for i in range(cols)]

    cell_h = min((avail_h - head_h) / len(lines) - row_gap, wide_w * FLAT_CELL)
    cell_h = max(cell_h, min_cell_h)
    # height left over on a big sheet goes into the gaps between rows, not into
    # the boxes: the type is limited by how wide a desk is, so a taller box only
    # buys white space inside itself, while a wider aisle is how the hall reads
    spare = (avail_h - head_h) - len(lines) * (cell_h + row_gap)
    if spare > 0:
        row_gap = min(row_gap + spare / len(lines), cell_h * 0.9)

    longest = max((len(r) for r in sub.Roll.astype(str)), default=9)
    roll_font, seat_font = _fit_font(cv, wide_w - 1.4, "B" * longest, cell_h)

    cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica-Bold", max(5.0, seat_font + 0.6))
    cv.drawRightString((x0 + gutter - 2) * mm, (y_top - head_h + 1.4) * mm, "COL")
    for i in range(cols):
        cv.setFont("Helvetica-Bold",
                   max(4.4, (seat_font + 0.6) if seated[i] else seat_font - 0.4))
        cv.drawCentredString((x_at[i] + (col_w[i] - 0.9) / 2) * mm,
                             (y_top - head_h + 1.4) * mm, str(i + 1))

    y = y_top - head_h - cell_h
    for k, row in lines:
        reserved = getattr(blocks[k], "reserved", set()) if blocks[k] is not None else set()
        cv.setFillColorRGB(*INK)
        cv.setFont("Helvetica-Bold", max(6.0, min(10.5, cell_h * 0.9)))
        cv.drawRightString((x0 + gutter - 2) * mm, (y + cell_h / 2 - 1.2) * mm,
                           f"Row {_row_no(k)}")
        _draw_seat_row(cv, row, key=k, x0=x0 + gutter, y=y, widths=col_w,
                       cell_h=cell_h, seat_roll=seat_roll, reserved=reserved,
                       roll_font=roll_font, seat_font=seat_font)
        y -= cell_h + row_gap
    return y + row_gap


INK, MUTE, RULE = (0.13, 0.13, 0.11), (0.42, 0.40, 0.35), (0.80, 0.78, 0.72)


def _qr_image(text: str):
    """QR as a reportlab-drawable image, or None if `qrcode` isn't installed -
    the poster still prints with the URL in large type, which is the part that
    actually has to work."""
    try:
        import qrcode
    except ImportError:
        return None
    img = qrcode.make(text, box_size=10, border=2)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    from reportlab.lib.utils import ImageReader
    return ImageReader(buf)


def hall_poster(out_pdf: str, *, room: str, url: str, cohort: str = "",
                course: str = C.COURSE, session: str = C.SESSION,
                counts: dict[str, int] | None = None, note: str = "") -> str:
    """One A4 poster for one hall door."""
    counts = counts or {}
    cv = canvas.Canvas(out_pdf, pagesize=(PAGE_W * mm, PAGE_H * mm))

    cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica-Bold", 9)
    cv.drawString(18 * mm, (PAGE_H - 20) * mm, f"{course.upper()}   ·   {session.upper()}")
    if C.COURSE_INSTRUCTORS:
        cv.setFont("Helvetica", 8.5)
        cv.drawString(18 * mm, (PAGE_H - 25) * mm, C.instructor_names())
    cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 30)
    cv.drawString(18 * mm, (PAGE_H - 33) * mm, "Find your block")
    cv.setFont("Helvetica", 14)
    label = room if " " in room else room.replace("LHC", "LHC-")
    cv.drawString(18 * mm, (PAGE_H - 43) * mm, f"{label}{('   ·   ' + cohort) if cohort else ''}")
    cv.setStrokeColorRGB(*INK); cv.setLineWidth(1.2)
    cv.line(18 * mm, (PAGE_H - 47) * mm, (PAGE_W - 18) * mm, (PAGE_H - 47) * mm)

    qr = _qr_image(url)
    qr_side = 62.0
    if qr:
        cv.drawImage(qr, (PAGE_W - 18 - qr_side) * mm, (PAGE_H - 47 - 8 - qr_side) * mm,
                     qr_side * mm, qr_side * mm, mask="auto")
        cv.setFont("Helvetica", 8); cv.setFillColorRGB(*MUTE)
        cv.drawCentredString((PAGE_W - 18 - qr_side / 2) * mm, (PAGE_H - 47 - 13 - qr_side) * mm,
                             "Scan with your phone camera")

    # The text column has to stop short of the QR, and the URL is the one line
    # that must never be truncated - shrink it to fit rather than clip it.
    col_w = PAGE_W - 36 - (qr_side + 8 if qr else 0)
    cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 13)
    cv.drawString(18 * mm, (PAGE_H - 62) * mm, "1.  Open this page")
    size = 13.0
    while size > 7 and cv.stringWidth(url, "Courier-Bold", size) > col_w * mm:
        size -= 0.5
    cv.setFont("Courier-Bold", size)
    cv.drawString(18 * mm, (PAGE_H - 71) * mm, url)
    cv.setFont("Helvetica-Bold", 13)
    cv.drawString(18 * mm, (PAGE_H - 85) * mm, "2.  Type your roll number")
    cv.setFont("Helvetica", 11); cv.setFillColorRGB(*MUTE)
    cv.drawString(18 * mm, (PAGE_H - 93) * mm, "It shows your block, highlighted on the hall map.")
    cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 13)
    cv.drawString(18 * mm, (PAGE_H - 107) * mm, "3.  Sit anywhere inside that block")

    # block strip - the colours students see on the plan and the sheets
    blocks = C.ROOMS.get(room, [])
    if blocks:
        y = (PAGE_H - 150)
        cv.setFont("Helvetica-Bold", 12); cv.setFillColorRGB(*INK)
        cv.drawString(18 * mm, (y + 20) * mm, "Blocks in this hall")
        w = (PAGE_W - 36) / max(len(blocks), 1)
        for i, b in enumerate(blocks):
            x = 18 + i * w
            col = C.BLOCK_COLOUR.get(b.key, "444444")
            rgb = tuple(int(col[j:j + 2], 16) / 255 for j in (0, 2, 4))
            n = counts.get(b.key, 0)
            cv.setFillColorRGB(*rgb) if n else cv.setFillColorRGB(*RULE)
            cv.rect(x * mm, y * mm, (w - 2) * mm, 17 * mm, stroke=0, fill=1)
            cv.setFillColorRGB(1, 1, 1); cv.setFont("Helvetica-Bold", 17)
            cv.drawCentredString((x + (w - 2) / 2) * mm, (y + 5.5) * mm, b.key)
            cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 7.5)
            cv.drawCentredString((x + (w - 2) / 2) * mm, (y - 5) * mm,
                                 f"{n} students" if n else "not used")

    cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 11)
    cv.drawString(18 * mm, (PAGE_H - 175) * mm, "Sit anywhere inside your own block.")
    cv.setFont("Helvetica", 10); cv.setFillColorRGB(*MUTE)
    cv.drawString(18 * mm, (PAGE_H - 183) * mm,
                  "Sign the attendance sheet for that block only - not a neighbour's.")

    cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 9)
    cv.drawString(18 * mm, 30 * mm, note or
                  "Blocks are fixed for the term. If your roll number is not found, check it "
                  "against your ID card, then ask a TA.")
    cv.setFont("Helvetica-Oblique", 8)
    cv.drawString(18 * mm, 24 * mm, "This page holds no names - only block allotments.")
    cv.showPage(); cv.save()
    return out_pdf


def roll_ranges(alloc, room_col: str = "Room") -> list[dict]:
    """One line per place a student can be sent: room and block, the roll
    numbers that belong there, and how many.

    Ranges are read off the *sorted* roll list, not the seating order: a
    student looks for their own roll number, and "B26CH1929 to B26CM1959" is
    something they can check at a glance. A group that is not one unbroken
    stretch (block B also holds the two students the roster gained at the end)
    prints each stretch, so nothing is quietly rounded over.

    A room numbered by row and column is one line for the whole room: there is
    no block to send anyone to.
    """
    rolls = sorted(str(r) for r in alloc.Roll)
    pos = {r: i for i, r in enumerate(rolls)}
    out = []
    for room in dict.fromkeys(alloc[room_col]) if room_col in alloc else [None]:
        sub = alloc[alloc[room_col] == room] if room_col in alloc else alloc
        if _is_row_col(room):
            groups = [("", sub)]
        else:
            groups = [(str(b), sub[sub.Block.astype(str) == str(b)])
                      for b in sorted({str(x) for x in sub.Block})]
        for block, g in groups:
            idx = sorted(pos[str(r)] for r in g.Roll)
            # break the list where the roll numbers stop running on, and again
            # where the branch changes: "B26AE1943 to B26BB1932" reads as one
            # sweep of roll numbers, but AE and BB are different branches and a
            # student checking their own number has to work out which side of
            # the join they are on
            runs, start, prev = [], idx[0], idx[0]
            for i in idx[1:]:
                if i != prev + 1 or _branch_of(rolls[i]) != _branch_of(rolls[prev]):
                    runs.append((start, prev))
                    start = i
                prev = i
            runs.append((start, prev))
            out.append({
                "room": str(room), "block": block, "n": len(g),
                "runs": [(rolls[a], rolls[b]) for a, b in runs],
                "rowcol": _is_row_col(room),
                "note": _batch_note(g, alloc),
            })
    return out


def _branch_of(roll: str) -> str:
    """`B26AE1943` -> `B26AE`: batch and branch, the part a roll number shares
    with its neighbours."""
    m = re.match(r"^([A-Za-z]\d{2}[A-Za-z]{2})", str(roll).strip())
    return m.group(1).upper() if m else str(roll)[:5].upper()


def _batch_note(group, alloc) -> str:
    """Say it in words where a range would mislead.

    The seniors are a handful of roll numbers from four earlier batches, and
    they all sit together. "B22AI914 to B26AE1901" is technically their range
    and tells a first-year nothing; "every roll number from B22, B23, B24 and
    B25, and B26AE1901" is what they are actually looking for.
    """
    def batch(roll):
        return str(roll)[:3].upper()

    counts = {}
    for r in alloc.Roll:
        counts[batch(r)] = counts.get(batch(r), 0) + 1
    current = max(counts, key=lambda b: counts[b])

    seniors_all = {str(r) for r in alloc.Roll if batch(r) != current}
    here = {str(r) for r in group.Roll}
    if not seniors_all or not seniors_all <= here:
        return ""
    batches = sorted({batch(r) for r in seniors_all})
    listed = ", ".join(batches[:-1]) + (" and " if len(batches) > 1 else "") + batches[-1]
    others = sorted(here - seniors_all)
    tail = ""
    if others:
        tail = (", and " + ", ".join(others) if len(others) <= 3
                else f", and {others[0]} to {others[-1]}")
    return f"every roll number from {listed}{tail}"


def _range_text(run) -> str:
    lo, hi = run
    return lo if lo == hi else f"{lo} to {hi}"


def report_summary(alloc, out_pdf: str, *, exam: str, course: str = C.COURSE,
                   when: str = "", room_col: str = "Room", report_by: str = "",
                   late_note: str = "", lookup_url: str = "") -> str:
    """The one sheet a student needs: find your roll number, go to that room.

    Everything else this module prints is for whoever is running the exam. This
    is for the noticeboard, the department group and the mail: one page, no
    seat numbers, no names.
    """
    rows = roll_ranges(alloc, room_col)
    cv = canvas.Canvas(out_pdf, pagesize=(PAGE_W * mm, PAGE_H * mm))

    cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica-Bold", 9)
    cv.drawString(15 * mm, (PAGE_H - 18) * mm, course.upper())
    if C.COURSE_INSTRUCTORS:
        cv.setFont("Helvetica", 8)
        cv.drawRightString((PAGE_W - 15) * mm, (PAGE_H - 18) * mm, C.instructor_names())
    cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 24)
    cv.drawString(15 * mm, (PAGE_H - 30) * mm, f"{exam} - where to report")
    cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 11)
    y = PAGE_H - 38
    if when:
        cv.drawString(15 * mm, y * mm, when)
        y -= 6
    cv.setFont("Helvetica", 10)
    cv.drawString(15 * mm, y * mm,
                  f"{len(alloc)} students. Find your roll number below and go to that "
                  "room and block.")
    y -= 4
    cv.setStrokeColorRGB(*INK); cv.setLineWidth(1.2)
    cv.line(15 * mm, y * mm, (PAGE_W - 15) * mm, y * mm)

    y -= 9
    cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica-Bold", 8)
    for x, lab in ((15, "ROOM"), (45, "BLOCK"), (68, "ROLL NUMBERS"), (PAGE_W - 15, "")):
        cv.drawString(x * mm, y * mm, lab)
    cv.drawRightString((PAGE_W - 15) * mm, y * mm, "STUDENTS")
    y -= 2
    cv.setStrokeColorRGB(*RULE); cv.setLineWidth(0.5)
    cv.line(15 * mm, y * mm, (PAGE_W - 15) * mm, y * mm)

    col_w = PAGE_W - 15 - 68 - 18
    for r in rows:
        text = ", ".join(_range_text(run) for run in r["runs"])
        lines = _wrap_text(cv, text, col_w, "Helvetica", 10)
        y -= 10
        cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 10.5)
        cv.drawString(15 * mm, y * mm, r["room"])
        cv.drawString(45 * mm, y * mm,
                      _group_label(r["room"], r["block"]).replace("Block ", "")
                      if r["block"] else "whole room")
        cv.setFont("Helvetica-Bold", 10.5)
        cv.drawRightString((PAGE_W - 15) * mm, y * mm, str(r["n"]))
        cv.setFillColorRGB(*INK); cv.setFont("Helvetica", 10)
        for i, line in enumerate(lines):
            if i:
                y -= 5.2
            cv.drawString(68 * mm, y * mm, line)
        if r["note"]:
            y -= 5
            cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica-Oblique", 9)
            cv.drawString(68 * mm, y * mm,
                          _fit_text(cv, r["note"], PAGE_W - 15 - 68 - 18, "Helvetica-Oblique", 9))
        if r["rowcol"]:
            y -= 5
            cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica-Oblique", 9)
            cv.drawString(68 * mm, y * mm,
                          "one hall, seats numbered by row and column - see the plan at the door")
        y -= 3
        cv.setStrokeColorRGB(*RULE); cv.setLineWidth(0.3)
        cv.line(15 * mm, y * mm, (PAGE_W - 15) * mm, y * mm)

    y -= 12
    cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 11)
    cv.drawString(15 * mm, y * mm, "On the day")
    notes = [n for n in (
        report_by or "Reach your room at least ten minutes before the start time.",
        "Carry your Institute ID card. Seats are checked against the seating list.",
        "Your own seat number is on the block sheet at the room door and on the "
        "seating plan inside.",
        late_note,
        f"Look yourself up: {lookup_url}" if lookup_url else "",
    ) if n]
    for note in notes:
        # wrapped, not truncated: a note cut off at "a TA will place ..." is
        # worse than no note
        for i, line in enumerate(_wrap_text(cv, note, PAGE_W - 34, "Helvetica", 9.5)):
            y -= 6
            cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 9.5)
            cv.drawString((15 if not i else 18) * mm, y * mm,
                          ("· " if not i else "") + line)

    cv.showPage()
    cv.save()
    return out_pdf


def exam_pack(alloc, out_pdf: str, *, room_map: dict, exam: str,
              course: str = C.COURSE, when: str = "", venue: str = "",
              spacing: str = "", duties=None, support=None, heads=None,
              room_col: str = "Room", landscape: bool = False,
              paper: str = "a4") -> str:
    """The printable exam pack: cover with the TA posting, a seat matrix per
    block drawn in the hall's real geometry, then a signature list per room.

    This replaces reading a seat plan off a screen. An invigilator gets the
    grid they are standing in front of, with roll numbers in position, and
    their own name on it; the office gets one file to print.

    `duties` is `ta_duty.by_block()` - {(room, block): [names]}; a room-level
    posting uses the key `(room, ta_duty.WHOLE_ROOM)`. `support` is
    `ta_duty.assign_support()` - the off-block jobs (question-paper printing),
    printed on the cover with names against them.
    """
    duties = duties or {}
    is_legal = str(paper).lower() == "legal"
    if is_legal:
        init_w, init_h = (LEGAL_W, LEGAL_H) if landscape else (LEGAL_H, LEGAL_W)
    else:
        init_w, init_h = (PAGE_H, PAGE_W) if landscape else (PAGE_W, PAGE_H)
    cv = canvas.Canvas(out_pdf, pagesize=(init_w * mm, init_h * mm))
    rooms = list(dict.fromkeys(alloc[room_col])) if room_col in alloc else ["-"]

    _cover(cv, alloc, exam=exam, course=course, when=when, venue=venue,
           spacing=spacing, duties=duties, rooms=rooms, room_col=room_col,
           support=support, heads=heads, landscape=landscape, paper=paper)

    serials = serial_numbers(alloc, room_col)
    for room in rooms:
        sub = alloc[alloc[room_col] == room] if room_col in alloc else alloc
        room_obj = room_map.get(room)
        _seat_pages(cv, sub, room, room_obj, exam=exam, course=course, when=when,
                    duties=duties, landscape=landscape, paper=paper)
        _signature_pages(cv, sub, room, room_obj, exam=exam, course=course, when=when,
                         duties=duties, serials=serials, landscape=landscape, paper=paper)
    cv.save()
    return out_pdf


# invigilator_pack is what the earlier version of this module called it; the
# dashboard and any saved notes still use that name.
def invigilator_pack(alloc, out_pdf: str, *, exam: str, course: str = C.COURSE,
                     when: str = "", venue: str = "", room_col: str = "Room",
                     room_map: dict | None = None, duties=None, spacing: str = "",
                     paper: str = "a4", landscape: bool = False) -> str:
    return exam_pack(alloc, out_pdf, room_map=room_map or {}, exam=exam, course=course,
                     when=when, venue=venue, spacing=spacing, duties=duties,
                     room_col=room_col, paper=paper, landscape=landscape)


def notice_board_map(alloc, out_pdf: str, *, room_map: dict, exam: str,
                     course: str = C.COURSE, when: str = "", venue: str = "",
                     room_col: str = "Room", paper: str = "auto",
                     paginate: bool | None = None) -> str:
    """One sheet per room for the board outside the hall: **A3 landscape for
    the big halls, A4 portrait for the rest**.

    Blocks are grouped by the hall's own sections where this project knows
    them (`config.ROOMS`), rear above front, each section a row of blocks side
    by side, so the sheet reads like standing at the door. A section too wide
    for the paper is broken into further bands rather than shrunk past
    legibility.
    """
    cv = canvas.Canvas(out_pdf, pagesize=(A3_W * mm, A3_H * mm))
    rooms = list(dict.fromkeys(alloc[room_col])) if room_col in alloc else ["-"]

    if str(paper).lower() not in {"auto", "a4", "legal"}:
        raise ValueError(f"paper must be 'auto', 'a4' or 'legal', not {paper!r}")
    a4_only = str(paper).lower() == "a4"
    is_legal = str(paper).lower() == "legal"

    for room in rooms:
        sub = alloc[alloc[room_col] == room] if room_col in alloc else alloc
        if is_legal:
            page_w, page_h = (LEGAL_W, LEGAL_H)
        elif a4_only:
            # A hall that needed A3 does not become readable by being scaled
            # onto A4 - it is split across sheets at the same cell size, to be
            # taped up side by side.
            page_w, page_h = ((PAGE_H, PAGE_W) if _wants_landscape(room, room_map.get(room), sub)
                              else (PAGE_W, PAGE_H))
        else:
            page_w, page_h = ((A3_W, A3_H) if str(room).strip() in A3_ROOMS
                              else (PAGE_W, PAGE_H))
        cv.setPageSize((page_w * mm, page_h * mm))
        paginate_room = (a4_only or is_legal) if paginate is None else paginate
        _notice_page(cv, sub, room, room_map.get(room), exam=exam, course=course,
                     when=when, venue=venue, page_w=page_w, page_h=page_h,
                     paginate=paginate_room)
    cv.save()
    return out_pdf


def _wants_landscape(room, room_obj, sub) -> bool:
    """A grid wider than it is tall wants the paper that way round."""
    if room_obj is not None and room_obj.blocks:
        widest = max((len(r) for b in room_obj.blocks.values() for r in b.rows), default=1)
    else:
        widest = int(sub.groupby("Block").size().max()) if len(sub) else 1
    return widest > LANDSCAPE_COLS


# Left-to-right seat order within each LHC 110 section, exactly as the
# seating-plan workbook and the uploaded HTML lay the hall out (stage at the
# top: "LEFT SIDE (Block D)", "CENTRE LEFT (Block C)", "CENTRE RIGHT (Block
# B)", "RIGHT SIDE (Block A)" for the front row; "LEFT SIDE (Block G)",
# "CENTRE (Block F)", "RIGHT SIDE (Block E)" for the rear) - alphabetical
# order (A,B,C,D / E,F,G) does not match how the room is actually arranged.
LHC110_BLOCK_ORDER = {"front": ["D", "C", "B", "A"], "rear": ["G", "F", "E"]}


def _ordered(keys, order) -> list[str]:
    """`keys` in `order`'s sequence, any key `order` doesn't know about
    tacked on the end sorted - so a block the workbook hasn't named yet still
    shows up instead of silently vanishing."""
    keys = set(keys)
    head = [k for k in order if k in keys]
    tail = sorted(keys - set(head))
    return head + tail


def _sections_for(room: str, blocks: dict) -> list[list[str]]:
    """Blocks grouped the way the hall is: rear gallery above front section
    where `config.ROOMS` says so, otherwise one row of blocks. Within a
    section, LHC 110 keeps the room's own left-to-right seat order rather
    than alphabetical."""
    known = C.ROOMS.get(str(room).replace(" ", ""))
    if not known:
        return [sorted(blocks)]
    section_of = {b.key: b.section for b in known}
    is_110 = str(room).strip() == "LHC 110"
    rear = [k for k in blocks if section_of.get(k) == "rear"]
    front = [k for k in blocks if section_of.get(k) != "rear"]
    if is_110:
        rear = _ordered(rear, LHC110_BLOCK_ORDER["rear"])
        front = _ordered(front, LHC110_BLOCK_ORDER["front"])
    else:
        rear, front = sorted(rear), sorted(front)
    return [g for g in (rear, front) if g]


def _notice_head(cv, sub, room, *, exam, course, when, page_w, page_h, margin,
                 cont: str = "") -> None:
    cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica-Bold", 10 if page_w > 300 else 8.5)
    cv.drawString(margin * mm, (page_h - 14) * mm, f"{course.upper()}  ·  {exam.upper()}")
    if C.COURSE_INSTRUCTORS:
        cv.setFont("Helvetica", 9 if page_w > 300 else 7.5)
        cv.drawRightString((page_w - margin) * mm, (page_h - 14) * mm, C.instructor_names())
    cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 22 if page_w > 300 else 17)
    cv.drawString(margin * mm, (page_h - 25) * mm, str(room) + cont)
    cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 11 if page_w > 300 else 9)
    tail = " · ".join(p for p in (when, f"{len(sub)} students",
                                 "seats by row and column" if _is_row_col(room) else "") if p)
    cv.drawRightString((page_w - margin) * mm, (page_h - 25) * mm, tail)
    cv.setStrokeColorRGB(*RULE); cv.setLineWidth(0.8)
    cv.line(margin * mm, (page_h - 30) * mm, (page_w - margin) * mm, (page_h - 30) * mm)


def _notice_page(cv, sub, room, room_obj, *, exam, course, when, venue,
                 page_w=A3_W, page_h=A3_H, paginate: bool = False) -> None:
    """One room on one sheet, unless `paginate` - then the room is broken over
    as many sheets as it needs at a cell size that stays readable, rather than
    shrunk until it fits. That is what A4-only printing means for a hall the
    size of LHC 110: four sheets to tape up, not one nobody can read."""
    import exam_rooms

    seat_roll = {(str(r.Block), str(r.Seat)): str(r.Roll) for r in sub.itertuples()}
    blocks = (dict(room_obj.blocks) if room_obj is not None
              else {b: None for b in dict.fromkeys(sub.Block)})
    blocks = {k: blocks[k] for k in sorted(blocks) if (sub.Block == k).any()}
    if not blocks:
        return

    margin = 12.0 if page_w > 300 else 10.0
    top, bottom = page_h - 34, 12.0
    if _is_row_col(room):
        _notice_head(cv, sub, room, exam=exam, course=course, when=when,
                     page_w=page_w, page_h=page_h, margin=margin)
        _rowcol_grid(cv, sub, room, blocks, x0=margin, y_top=top,
                     avail_w=page_w - 2 * margin, avail_h=top - bottom,
                     seat_roll=seat_roll, row_gap=1.4)
        cv.showPage()
        return
    label_h, block_gap, section_gap, row_gap = 6.0, 6.0, 8.0, 1.0
    # LHC 110 gets a shared row-number gutter down the left (one hall row is
    # the same physical row across every block that sits in it) and a column
    # header under each block's own label - everywhere else this stays 0 and
    # the sheet draws exactly as it always has.
    axis_gutter = 8.0 if str(room).strip() == "LHC 110" else 0.0
    axis_h = 4.5 if axis_gutter else 0.0
    width_budget = page_w - 2 * margin - axis_gutter
    grid_x0 = margin + axis_gutter
    # Below this a used desk is too narrow for a roll number read from across a
    # corridor. Empty desks print narrow (NARROW_COL), so this is the width of
    # the columns people are actually sitting in - the ones the sheet is for.
    min_cell_w = 15.0 if page_w > 400 else (13.5 if page_w > 300 else 11.0)

    longest = max((len(r) for r in sub.Roll.astype(str)), default=9)
    probe = cv.stringWidth("B" * longest, "Helvetica-Bold", 6.0) / mm + 2.5
    grids = {}
    for k, blk in blocks.items():
        raw = blk.rows if blk is not None else [
            sorted(sub[sub.Block == k].Seat,
                   key=lambda s: int(re.sub(r"\D", "", str(s)) or 0))]
        grids[k] = exam_rooms._wrap_rows(raw, width=max(4, int(width_budget // probe)))

    # every second column of a spaced block is empty; those print narrow and
    # give their width to the columns people are sitting in
    seated_by = {k: _seated_cols(g, k, seat_roll) for k, g in grids.items()}

    def weight(k):
        return _col_weight(seated_by[k])

    def band_w(k, wide):
        return sum(_col_widths(seated_by[k], wide))

    # A section wider than the paper is split into further bands: shrinking to
    # fit would put five blocks of a small room in 6mm cells.
    sections = []
    for group in _sections_for(room, grids):
        band = []
        for k in [k for k in group if k in grids]:
            trial = band + [k]
            need = sum(weight(x) for x in trial) * min_cell_w + (len(trial) - 1) * block_gap
            if band and need > width_budget:
                sections.append(band)
                band = [k]
            else:
                band = trial
        if band:
            sections.append(band)
    sections = [g for g in sections if g]

    cell_w = min((width_budget - (len(g) - 1) * block_gap) / max(sum(weight(k) for k in g), 1)
                 for g in sections)
    rows_stack = sum(max(len(grids[k]) for k in g) for g in sections)
    height_budget = (top - bottom) - (label_h + axis_h) * len(sections) \
        - section_gap * max(len(sections) - 1, 0) - row_gap * rows_stack
    # a wide, shallow room (ten blocks of one row) leaves height to spare, so
    # let the cells grow past square rather than printing a band of stamps
    # across the top of an empty sheet
    if paginate:
        tallest = max(max(len(grids[k]) for k in g) for g in sections)
        page_h_budget = (top - bottom) - label_h - section_gap
        cell_h = min(cell_w * FLAT_CELL, page_h_budget / max(tallest, 1) - row_gap)
        # the flat box leaves height over on the fullest sheet: into the aisles,
        # so the tallest band fills its page instead of stopping two thirds down
        spare = page_h_budget - tallest * (cell_h + row_gap)
        if spare > 0:
            row_gap = min(row_gap + spare / tallest, cell_h * 0.9)
    else:
        cell_h = min(height_budget / max(rows_stack, 1), cell_w * FLAT_CELL)
        leftover = height_budget - cell_h * rows_stack
        # a flat box leaves height over: it goes into the aisles between rows
        # and between sections, since a taller box would only be white space
        if leftover > 0:
            row_gap = min(row_gap + leftover / (2 * rows_stack), cell_h * 0.6)
            leftover = height_budget - cell_h * rows_stack - (row_gap - 1.0) * rows_stack
        if len(sections) > 1 and leftover > 0:
            section_gap = min(34.0, section_gap + leftover / (len(sections) - 1))

    roll_font, seat_font = _fit_font(cv, cell_w - 1.4, "B" * longest, cell_h)

    def band_h(group):
        return label_h + axis_h + max(len(grids[k]) for k in group) * (cell_h + row_gap)

    # count the sheets before drawing any, so sheet 1 can say what it is one of:
    # a wall map taped up out of order is worse than no map
    total_sheets, y_probe = 1, top
    if paginate:
        for group in sections:
            if y_probe - band_h(group) < bottom and y_probe < top:
                total_sheets += 1
                y_probe = top
            y_probe -= band_h(group) + section_gap

    def head(sheet: int):
        _notice_head(cv, sub, room, exam=exam, course=course, when=when,
                     page_w=page_w, page_h=page_h, margin=margin,
                     cont=f"  (sheet {sheet} of {total_sheets})" if total_sheets > 1 else "")

    head(1)

    y_section = top
    sheet = 1
    for group in sections:
        if paginate and y_section - band_h(group) < bottom and y_section < top:
            cv.showPage()
            sheet += 1
            head(sheet)
            y_section = top
        cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 9)
        x = grid_x0
        for k in group:
            cv.drawString(x * mm, (y_section - 4) * mm,
                          f"{_group_label(room, k)}  ({int((sub.Block == k).sum())})")
            x += band_w(k, cell_w) + block_gap
        y_top = y_section - label_h

        if axis_gutter:
            cv.setFillColorRGB(*MUTE)
            cv.drawRightString((grid_x0 - 2) * mm, (y_top - axis_h + 3.0) * mm, "COL")
            x = grid_x0
            for k in group:
                widths = _col_widths(seated_by[k], cell_w)
                labels = _lhc110_columns(room, k) or []
                for i in range(len(widths)):
                    lbl = labels[i] if i < len(labels) else ""
                    cv.setFont("Helvetica-Bold", max(4.4, (seat_font * 0.9) if seated_by[k][i] else seat_font * 0.75))
                    cv.drawCentredString((x + sum(widths[:i]) + (widths[i] - 0.9) / 2) * mm,
                                         (y_top - axis_h + 3.0) * mm, str(lbl))
                x += band_w(k, cell_w) + block_gap
            # one shared "Row n" gutter for the whole section - a hall row is
            # the same physical row under every block that sits in it, so the
            # numbers only need drawing once, against whichever block in this
            # group actually has the most of them
            tallest = max((grids[k] for k in group), key=len)
            cv.setFillColorRGB(*INK)
            cv.setFont("Helvetica-Bold", max(6.0, min(9.0, cell_h * 0.85)))
            yy = y_top - axis_h - cell_h
            for r_i in range(len(tallest)):
                cv.drawRightString((grid_x0 - 2) * mm, (yy + cell_h / 2 - 1.0) * mm,
                                   f"Row {r_i + 1}")
                yy -= cell_h + row_gap
            y_top -= axis_h

        x = grid_x0
        for k in group:
            reserved = getattr(blocks[k], "reserved", set()) if blocks[k] is not None else set()
            widths = _col_widths(seated_by[k], cell_w)
            y = y_top - cell_h
            for row in grids[k]:
                _draw_seat_row(cv, row, key=k, x0=x, y=y, widths=widths,
                               cell_h=cell_h, seat_roll=seat_roll, reserved=reserved,
                               roll_font=roll_font, seat_font=seat_font)
                y -= cell_h + row_gap
            x += band_w(k, cell_w) + block_gap
        y_section = y_top - max(len(grids[k]) for k in group) * (cell_h + row_gap) \
            - section_gap
    cv.showPage()


def _cover_landscape(cv, alloc, *, exam, course, when, venue, spacing, duties, rooms, room_col,
                     support=None, heads=None, paper: str = "a4") -> None:
    is_legal = str(paper).lower() == "legal"
    pw, ph = (LEGAL_W, LEGAL_H) if is_legal else (PAGE_H, PAGE_W)  # 355.6, 215.9 vs 297, 210
    cv.setPageSize((pw * mm, ph * mm))

    # Top Header
    cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica-Bold", 8.5)
    cv.drawString(15 * mm, (ph - 15) * mm, course.upper())
    if C.COURSE_INSTRUCTORS:
        cv.setFont("Helvetica", 8)
        cv.drawRightString((pw - 15) * mm, (ph - 15) * mm, C.instructor_names())

    cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 20)
    cv.drawString(15 * mm, (ph - 24) * mm, exam)

    cv.setFont("Helvetica", 9.5); cv.setFillColorRGB(*MUTE)
    sub_parts = [p for p in (when or "date and time to be announced",
                             f"Venue: {venue}" if venue else "",
                             f"Seating: {spacing}" if spacing else "") if p]
    cv.drawString(15 * mm, (ph - 31) * mm, "   ·   ".join(sub_parts))

    y_rule = ph - 35
    cv.setStrokeColorRGB(*INK); cv.setLineWidth(1.0)
    cv.line(15 * mm, y_rule * mm, (pw - 15) * mm, y_rule * mm)

    # 2 columns below header
    col1_x = 15
    col1_w = 160 if is_legal else 138
    y = y_rule - 9
    cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 11)
    cv.drawString(col1_x * mm, y * mm, "Invigilation")
    y -= 6
    cv.setFont("Helvetica-Bold", 7.5); cv.setFillColorRGB(*MUTE)
    stud_x = 44 if is_legal else 40
    inv_x = 64 if is_legal else 60
    for x_off, lab in ((0, "Room"), (22, "Block"), (stud_x, "Students"), (inv_x, "Invigilators")):
        cv.drawString((col1_x + x_off) * mm, y * mm, lab.upper())
    y -= 2
    cv.setStrokeColorRGB(*RULE); cv.setLineWidth(0.5)
    cv.line(col1_x * mm, y * mm, (col1_x + col1_w) * mm, y * mm)

    total_tas = 0
    for room in rooms:
        sub = alloc[alloc[room_col] == room] if room_col in alloc else alloc
        keys = [k for k in duties if k[0] == room] or [(room, "-")]
        for _, block in sorted(keys, key=lambda k: str(k[1])):
            names = duties.get((room, block), [])
            total_tas += len(names)
            n = int((sub.Block == block).sum()) if block != "-" else len(sub)
            y -= 6.2
            cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 8)
            cv.drawString(col1_x * mm, y * mm, str(room))
            cv.drawString((col1_x + 22) * mm, y * mm,
                          str(block) if block == "-" else _group_label(room, block))
            cv.setFont("Helvetica", 8)
            cv.drawString((col1_x + stud_x) * mm, y * mm, str(n))
            cv.setFillColorRGB(*(INK if names else (0.65, 0.25, 0.15)))
            cv.drawString((col1_x + inv_x) * mm, y * mm,
                          _fit_text(cv, ", ".join(names) or "- not assigned -",
                                    col1_w - inv_x, "Helvetica", 8))
            cv.setStrokeColorRGB(*RULE); cv.setLineWidth(0.3)
            cv.line(col1_x * mm, (y - 2) * mm, (col1_x + col1_w) * mm, (y - 2) * mm)

    # Column 2: Other duties & Exam management
    col2_x = 185 if is_legal else 160
    col2_w = pw - 15 - col2_x
    y2 = y_rule - 9

    if support is not None and len(support):
        cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 11)
        cv.drawString(col2_x * mm, y2 * mm, "Other duties")
        for role in dict.fromkeys(support.Role):
            names = [r.TA for r in support.itertuples() if r.Role == role and r.TA]
            y2 -= 6.0
            cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 8)
            cv.drawString(col2_x * mm, y2 * mm, str(role))
            cv.setFont("Helvetica", 8)
            cv.drawString((col2_x + 40) * mm, y2 * mm,
                          _fit_text(cv, ", ".join(names) or "- not assigned -",
                                    col2_w - 40, "Helvetica", 8))
        y2 -= 6

    if heads is not None and len(heads):
        cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 11)
        cv.drawString(col2_x * mm, y2 * mm, "Exam management")
        for duty in dict.fromkeys(heads.Duty):
            rows_h = heads[heads.Duty == duty]
            lead = bool(rows_h.InCharge.any())
            y2 -= 5.8
            cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 8)
            cv.drawString(col2_x * mm, y2 * mm, f"{rows_h.iloc[0].TA}")
            cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 7.5)
            cv.drawString((col2_x + 40) * mm, y2 * mm, duty + (" · IN CHARGE" if lead else ""))
            for job in rows_h.Responsibility:
                y2 -= 4.8
                cv.drawString((col2_x + 5) * mm, y2 * mm, "· " + _fit_text(cv, str(job), col2_w - 5, "Helvetica", 7.5))
            y2 -= 2

    # Footer
    cv.setStrokeColorRGB(*RULE); cv.setLineWidth(0.6)
    cv.line(15 * mm, 32 * mm, (pw - 15) * mm, 32 * mm)

    cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 8)
    pack_format = "landscape Legal (14×8.5 in)" if is_legal else "landscape"
    cv.drawString(15 * mm, 25.5 * mm,
                  f"{len(alloc)} students · {len(rooms)} room(s) · {total_tas} invigilators · All pages in this pack are {pack_format}.")
    cv.drawString(15 * mm, 19.5 * mm,
                  "One seat grid per block, followed by signature sheets per block; each invigilator holds only their own.")

    # Room-specific instructions: where the seat number actually is, and how
    # a room with no printed numbers gets filled correctly anyway.
    cv.setFont("Helvetica-Oblique", 8)
    cv.drawString(15 * mm, 13.5 * mm,
                  "LHC 110: the seat number is marked on the back of the seat; look behind you if it isn't visible from the front.")
    cv.drawString(15 * mm, 7.5 * mm,
                  "LHC 105: seats are not numbered; TAs seat students according to this plan.")

    cv.showPage()


def _cover(cv, alloc, *, exam, course, when, venue, spacing, duties, rooms, room_col,
           support=None, heads=None, landscape: bool = False, paper: str = "a4") -> None:
    if landscape:
        return _cover_landscape(cv, alloc, exam=exam, course=course, when=when, venue=venue,
                                spacing=spacing, duties=duties, rooms=rooms, room_col=room_col,
                                support=support, heads=heads, paper=paper)
    is_legal = str(paper).lower() == "legal"
    pw, ph = (LEGAL_H, LEGAL_W) if is_legal else (PAGE_W, PAGE_H)
    cv.setPageSize((pw * mm, ph * mm))
    cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica-Bold", 9)
    cv.drawString(15 * mm, (ph - 18) * mm, course.upper())
    cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 24)
    cv.drawString(15 * mm, (ph - 30) * mm, exam)
    cv.setFont("Helvetica", 11); cv.setFillColorRGB(*MUTE)
    cv.drawString(15 * mm, (ph - 38) * mm, when or "date and time to be announced")
    y_rule = ph - 48
    if spacing:
        cv.drawString(15 * mm, (ph - 44) * mm, f"Seating: {spacing}")
    if C.COURSE_INSTRUCTORS:
        cv.setFont("Helvetica", 9.5)
        cv.drawString(15 * mm, (ph - 50) * mm,
                      f"Course instructors: {C.instructor_names()}")
        y_rule = ph - 54
    cv.setStrokeColorRGB(*INK); cv.setLineWidth(1.2)
    cv.line(15 * mm, y_rule * mm, (pw - 15) * mm, y_rule * mm)

    y = y_rule - 12
    cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 12)
    cv.drawString(15 * mm, y * mm, "Invigilation")
    y -= 8
    cv.setFont("Helvetica-Bold", 8); cv.setFillColorRGB(*MUTE)
    for x, lab in ((15, "Room"), (42, "Block"), (60, "Students"), (85, "Invigilators")):
        cv.drawString(x * mm, y * mm, lab.upper())
    y -= 2
    cv.setStrokeColorRGB(*RULE); cv.setLineWidth(0.5)
    cv.line(15 * mm, y * mm, (pw - 15) * mm, y * mm)

    total_tas = 0
    for room in rooms:
        sub = alloc[alloc[room_col] == room]
        keys = [k for k in duties if k[0] == room] or [(room, "-")]
        for _, block in sorted(keys, key=lambda k: str(k[1])):
            names = duties.get((room, block), [])
            total_tas += len(names)
            n = int((sub.Block == block).sum()) if block != "-" else len(sub)
            y -= 7
            if y < 40:
                cv.showPage(); y = ph - 30
            cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 9)
            cv.drawString(15 * mm, y * mm, str(room))
            cv.drawString(42 * mm, y * mm,
                          str(block) if block == "-" else _group_label(room, block))
            cv.setFont("Helvetica", 9)
            cv.drawString(60 * mm, y * mm, str(n))
            cv.setFillColorRGB(*(INK if names else (0.65, 0.25, 0.15)))
            cv.drawString(85 * mm, y * mm,
                          _fit_text(cv, ", ".join(names) or "- not assigned -",
                                    pw - 15 - 85, "Helvetica", 9))
            cv.setStrokeColorRGB(*RULE); cv.setLineWidth(0.3)
            cv.line(15 * mm, (y - 2.5) * mm, (pw - 15) * mm, (y - 2.5) * mm)

    # off-block duties: named, because "someone will print the papers" is how
    # a hall ends up with 364 students and 340 papers
    if support is not None and len(support):
        y -= 14
        cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 12)
        cv.drawString(15 * mm, y * mm, "Other duties")
        for role in dict.fromkeys(support.Role):
            names = [r.TA for r in support.itertuples() if r.Role == role and r.TA]
            also = [r.TA for r in support.itertuples()
                    if r.Role == role and r.TA and getattr(r, "AlsoInvigilating", False)]
            y -= 7
            if y < 40:
                cv.showPage(); y = ph - 30
            cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 9)
            cv.drawString(15 * mm, y * mm, str(role))
            cv.setFillColorRGB(*(INK if names else (0.65, 0.25, 0.15)))
            cv.setFont("Helvetica", 9)
            note = " (also invigilating)" if also else ""
            cv.drawString(85 * mm, y * mm,
                          _fit_text(cv, (", ".join(names) or "- not assigned -") + note,
                                    pw - 15 - 85, "Helvetica", 9))

    # who is running the exam, and what each head owns
    if heads is not None and len(heads):
        y -= 14
        if y < 60:
            cv.showPage(); y = ph - 30
        cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 12)
        cv.drawString(15 * mm, y * mm, "Exam management")
        for duty in dict.fromkeys(heads.Duty):
            rows_h = heads[heads.Duty == duty]
            lead = bool(rows_h.InCharge.any())
            y -= 7
            if y < 30:
                cv.showPage(); y = ph - 30
            cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 9)
            cv.drawString(15 * mm, y * mm, f"{rows_h.iloc[0].TA}")
            cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 8)
            cv.drawString(85 * mm, y * mm, duty + ("  ·  OVERALL IN CHARGE" if lead else ""))
            for job in rows_h.Responsibility:
                y -= 5.5
                if y < 25:
                    cv.showPage(); y = ph - 30
                cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 8.5)
                cv.drawString(20 * mm, y * mm, "· " + _fit_text(cv, str(job), pw - 40,
                                                                "Helvetica", 8.5))
            y -= 2

    y -= 14
    cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 9)
    cv.drawString(15 * mm, y * mm,
                  f"{len(alloc)} students · {len(rooms)} room(s) · {total_tas} invigilators")
    y -= 6
    for line in ("One seat grid per block, then one signature sheet per block - each "
                 "invigilator holds only their own.",
                 "A room numbered by row and column prints as one grid, rows down the "
                 "side and columns across the top.",
                 "A list longer than one page prints two columns to a landscape sheet."):
        cv.drawString(15 * mm, y * mm, _fit_text(cv, line, pw - 30, "Helvetica", 9))
        y -= 5
    cv.showPage()


def _wrap_text(cv, text: str, width_mm: float, font: str, size: float) -> list[str]:
    """Break on spaces to fit the width. Used where the text has room to run
    on rather than be cut short."""
    words, lines, line = str(text).split(), [], ""
    for w in words:
        trial = (line + " " + w).strip()
        if line and cv.stringWidth(trial, font, size) > width_mm * mm:
            lines.append(line)
            line = w
        else:
            line = trial
    if line:
        lines.append(line)
    return lines or [""]


def _fit_text(cv, text: str, width_mm: float, font: str, size: float) -> str:
    if cv.stringWidth(text, font, size) <= width_mm * mm:
        return text
    while text and cv.stringWidth(text + "…", font, size) > width_mm * mm:
        text = text[:-1]
    return text + "…"


def _seat_pages(cv, sub, room, room_obj, *, exam, course, when, duties,
                landscape: bool = False, paper: str = "a4") -> None:
    """Seat grids: per block for the big halls, one page for a small room.

    The grid is sized to fill the page rather than sit in the top third. Cell
    width is the usable width divided by the widest row; cell height takes the
    space that is actually left, capped by an aspect ratio so a two-row block
    does not print as a wall of letterboxes. Type scales with the cell, since a
    6pt roll number in a 25mm box is the worst of both.

    Columns still bend to the content: a block the exam sheet gave as one flat
    row (LHC 206/207 arrive as a single 15-seat line) is re-wrapped narrow
    enough that a roll number fits. A block with real geometry is never
    re-wrapped, since its shape is the hall's.
    """
    import exam_rooms

    seat_roll = {(str(r.Block), str(r.Seat)): str(r.Roll) for r in sub.itertuples()}
    blocks = (dict(room_obj.blocks) if room_obj is not None
              else {b: None for b in dict.fromkeys(sub.Block)})
    blocks = {k: blocks[k] for k in sorted(blocks) if (sub.Block == k).any()}
    if not blocks:
        return

    longest = max((len(r) for r in sub.Roll.astype(str)), default=9)
    is_legal = str(paper).lower() == "legal"
    if is_legal:
        pw, ph = (LEGAL_W, LEGAL_H) if landscape else (LEGAL_H, LEGAL_W)
    else:
        pw, ph = (PAGE_H, PAGE_W) if landscape else (PAGE_W, PAGE_H)
    usable_w = pw - 30
    probe = cv.stringWidth("B" * longest, "Helvetica-Bold", 6.4) / mm + 3.0
    max_cols = max(4, int(usable_w // probe))

    def grid_for(block_key, block):
        raw = block.rows if block is not None else [
            sorted(sub[sub.Block == block_key].Seat,
                   key=lambda s: int(re.sub(r"\D", "", str(s)) or 0))]
        return exam_rooms._wrap_rows(raw, width=max_cols)

    def fonts_for(wide_w, cell_h):
        """Type that grows with the box, measured against the widest roll."""
        return _fit_font(cv, wide_w - 1.4, "B" * longest, cell_h)

    def draw_rows(rows, block_key, reserved, y, widths, cell_h, roll_font, seat_font,
                  gap=2.0, x0=15, row0=0):
        for i, row in enumerate(rows):
            # always left-aligned: centring a short last row would put seat 11
            # under seat 3, which is not where it is in the hall
            _draw_seat_row(cv, row, key=block_key, x0=x0, y=y, widths=widths,
                           cell_h=cell_h, seat_roll=seat_roll, reserved=reserved,
                           roll_font=roll_font, seat_font=seat_font)
            if x0 > 15:
                # a row-number gutter was reserved for this block (LHC 110) -
                # row0 carries the count across a block split onto a
                # "(cont.)" page, so the numbering never restarts at 1
                cv.setFillColorRGB(*INK)
                cv.setFont("Helvetica-Bold", max(6.0, min(9.5, cell_h * 0.85)))
                cv.drawRightString((x0 - 2) * mm, (y + cell_h / 2 - 1.1) * mm,
                                   f"Row {row0 + i + 1}")
            y -= cell_h + gap
        # hand back the bottom of the *last* row, not one step past it: the
        # caller adds its own spacing, and counting the step twice is what
        # pushed a five-block room onto a second sheet
        return y + cell_h + gap

    top = ph - 50 if landscape else ph - 56
    bottom = 14 if landscape else 18

    if _is_block_wise(room):
        cv.setPageSize((pw * mm, ph * mm))
        for block_key, block in blocks.items():
            in_block = sub[sub.Block == block_key]
            rows = grid_for(block_key, block)
            reserved = getattr(block, "reserved", set()) if block is not None else set()
            gap = 2.0
            col_labels = _lhc110_columns(room, block_key)
            # a row-number gutter and a column-number header, only where the
            # source workbook actually gives column numbers (LHC 110) - every
            # other block-wise hall keeps its old, unlabelled grid
            gutter = 10.0 if col_labels else 0.0
            head_h = 5.0 if col_labels else 0.0
            grid_x0 = 15 + gutter
            # a little slack, so a rounding error does not push the last row of
            # a fourteen-row block onto a second sheet
            avail_h = top - bottom - 6 - head_h
            # every second column of a spaced block is empty: those print narrow
            # and hand their width to the columns people are actually sitting in
            seated = _seated_cols(rows, block_key, seat_roll)
            wide_w = (usable_w - gutter) / _col_weight(seated)
            widths = _col_widths(seated, wide_w)
            # flat, not square: the box holds one short line of type, and the
            # height past that only pushes the roll number down the page
            cell_h = min(avail_h / max(len(rows), 1) - gap, wide_w * FLAT_CELL)
            roll_font, seat_font = fonts_for(wide_w, cell_h)
            names = duties.get((room, block_key)) or duties.get((room, "-")) or []

            # +0.5 so a row that fits exactly is not rounded onto a second page
            per_page = max(1, int((avail_h + 0.5) // (cell_h + gap)))
            if len(rows) <= per_page:
                # the block fits: spend what is left on the aisles between rows
                spare = avail_h - len(rows) * (cell_h + gap)
                if spare > 0:
                    gap = min(gap + spare / len(rows), cell_h * 0.9)
            for page_no, start in enumerate(range(0, len(rows), per_page)):
                chunk = rows[start:start + per_page]
                _pack_header(cv, f"{course} · {exam}",
                             f"{room} - Block {block_key}" + (" (cont.)" if page_no else ""),
                             when, "", len(in_block), tas=", ".join(names),
                             page_w=pw, page_h=ph)
                # y is a cell's *bottom* edge, so the first row starts a full
                # cell height below the header rule
                y_head = top - head_h
                if col_labels:
                    cv.setFillColorRGB(*MUTE)
                    cv.setFont("Helvetica-Bold", max(5.0, seat_font + 0.4))
                    cv.drawRightString((grid_x0 - 2) * mm, (y_head + 1.4) * mm, "COL")
                    for i in range(len(widths)):
                        lbl = col_labels[i] if i < len(col_labels) else ""
                        cv.setFont("Helvetica-Bold",
                                   max(4.4, (seat_font + 0.4) if seated[i] else seat_font - 0.6))
                        cv.drawCentredString(
                            (grid_x0 + sum(widths[:i]) + (widths[i] - 0.9) / 2) * mm,
                            (y_head + 1.4) * mm, str(lbl))
                y = y_head - cell_h - 2
                draw_rows(chunk, block_key, reserved, y, widths, cell_h,
                          roll_font, seat_font, gap=gap, x0=grid_x0, row0=start)
                cv.showPage()
        if not landscape:
            cv.setPageSize((pw * mm, ph * mm))
        return

    if _is_row_col(room):
        # fifteen columns across a portrait A4 is an 11mm desk and a 5pt roll
        # number: legible on screen, not across a hall. The same hall turned
        # sideways gives an 18mm desk and 8pt type on the same sheet of paper,
        # so a wide room prints landscape.
        cols_wide = max((len(r) for b in blocks.values() if b is not None
                         for r in b.rows), default=1)
        land = True if landscape else (cols_wide > LANDSCAPE_COLS)
        pw_rc, ph_rc = (LEGAL_W, LEGAL_H) if is_legal else ((PAGE_H, PAGE_W) if land else (PAGE_W, PAGE_H))
        cv.setPageSize((pw_rc * mm, ph_rc * mm))
        names = duties.get((room, "-")) or []
        _pack_header(cv, f"{course} · {exam}", f"{room} - seating", when, "", len(sub),
                     tas=", ".join(names), page_w=pw_rc, page_h=ph_rc)
        _rowcol_grid(cv, sub, room, blocks, x0=15, y_top=ph_rc - (50 if landscape else 56),
                     avail_w=pw_rc - 30, avail_h=ph_rc - (50 if landscape else 56) - bottom, seat_roll=seat_roll)
        cv.showPage()
        if not landscape:
            cv.setPageSize((pw * mm, ph * mm))     # back to portrait for the rest
        return

    # ── room-wise: the whole room on one page, blocks labelled inside ──────
    grids = {k: grid_for(k, v) for k, v in blocks.items()}
    total_rows = sum(len(g) for g in grids.values())
    # one width unit for the whole room, so every block reads at the same scale
    seated_by = {k: _seated_cols(g, k, seat_roll) for k, g in grids.items()}
    wide_w = min((usable_w / _col_weight(sc) for sc in seated_by.values()),
                 default=usable_w)
    widths_by = {k: _col_widths(sc, wide_w) for k, sc in seated_by.items()}
    label_h = 6.5
    gap = 1.5
    block_gap = 3.5          # breathing room between one block and the next
    # Solve for the cell height that makes the whole room fit one page:
    #   H = blocks*(label + block_gap) + rows*cell_h + (rows - blocks)*gap
    # Counting only the rows, as an earlier version did, left every block
    # short by its own label and pushed a five-block room onto a second sheet.
    n_blocks = max(len(grids), 1)
    budget = (top - bottom - 4) - n_blocks * (label_h + block_gap) \
        - (total_rows - n_blocks) * gap
    cell_h = min(budget / max(total_rows, 1), wide_w * FLAT_CELL)
    # a flat box leaves height over: give it to the space between rows and,
    # more importantly, between one block and the next - a block label sitting
    # on the last row of the block above reads as part of it
    spare = budget - cell_h * total_rows
    if spare > 0:
        gap = min(gap + spare / (2 * total_rows), cell_h * 0.6)
        block_gap = min(block_gap + spare / (2 * n_blocks), 14.0)
    roll_font, seat_font = fonts_for(wide_w, cell_h)
    names = duties.get((room, "-")) or []

    _pack_header(cv, f"{course} · {exam}", f"{room} - seating", when, "", len(sub),
                 tas=", ".join(names), page_w=pw, page_h=ph)
    y = top - 2
    for block_key, rows in grids.items():
        block = blocks[block_key]
        reserved = getattr(block, "reserved", set()) if block is not None else set()
        needed_h = label_h + cell_h + (len(rows) - 1) * (cell_h + gap) + block_gap
        if y - needed_h < bottom:
            cv.showPage()
            _pack_header(cv, f"{course} · {exam}", f"{room} - seating (cont.)", when, "",
                         len(sub), tas=", ".join(names), page_w=pw, page_h=ph)
            y = top
        cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 10)
        cv.drawString(15 * mm, y * mm, _group_label(room, block_key))
        cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 8)
        cv.drawString(34 * mm, y * mm,
                      f"· {int((sub.Block == block_key).sum())} students")
        y -= (label_h + cell_h)
        y = draw_rows(rows, block_key, reserved, y, widths_by[block_key], cell_h,
                      roll_font, seat_font, gap=gap)
        y -= block_gap
    cv.showPage()


def _signature_sort(sub, room):
    """The order a signature sheet reads in: block by block (row by row in a
    room numbered that way), then up the seats."""
    def num(v):
        return int(re.sub(r"\D", "", str(v)) or 0)

    row_col = _is_row_col(room)
    return sub.sort_values(
        ["Block", "Seat"],
        key=lambda c: c.map(num) if (c.name == "Seat" or row_col) else c)


def serial_numbers(alloc, room_col: str = "Room") -> dict:
    """One running number over the whole exam, in the order the signature
    sheets print: room by room, block by block, seat by seat.

    It runs on across rooms on purpose. The heads collect one pile of sheets at
    the end, and a serial that restarts in every hall makes "number 47" a
    question rather than an answer.
    """
    out, n = {}, 0
    rooms = list(dict.fromkeys(alloc[room_col])) if room_col in alloc else [None]
    for room in rooms:
        sub = alloc[alloc[room_col] == room] if room_col in alloc else alloc
        for r in _signature_sort(sub, room).itertuples():
            n += 1
            out[str(r.Roll)] = n
    return out


def _signature_pages(cv, sub, room, room_obj, *, exam, course, when, duties,
                     serials=None, landscape: bool = False, paper: str = "a4") -> None:
    """One signature sheet per block in the big halls, one per room elsewhere.

    A room-wide list in LHC 110 would mean the invigilator standing in block E
    holds a sheet whose first two pages belong to blocks A and B. In a small
    room the opposite is true: two TAs cover the whole room, and five sheets of
    eight students is five things to lose rather than one to hand over. The
    block is still printed against every row either way.
    """
    is_legal = str(paper).lower() == "legal"
    if is_legal:
        pw_sig, ph_sig = (LEGAL_W, LEGAL_H) if landscape else (LEGAL_H, LEGAL_W)
    else:
        pw_sig, ph_sig = (PAGE_H, PAGE_W) if landscape else (PAGE_W, PAGE_H)
    cv.setPageSize((pw_sig * mm, ph_sig * mm))
    block_wise = _is_block_wise(room)
    row_col = _is_row_col(room)
    blocks = sorted({str(b) for b in sub.Block})
    groups = ([(b, sub[sub.Block == b]) for b in blocks] if block_wise
              else [(None, sub)])

    serials = serials if serials is not None else serial_numbers(sub, "Room")
    for block_key, rows_df in groups:
        rows_df = _signature_sort(rows_df, room)
        if rows_df.empty:
            continue
        if block_wise:
            title = f"{room} - {_group_label(room, block_key)} · signatures"
            names = ", ".join(duties.get((room, block_key)) or duties.get((room, "-")) or [])
            tally = _group_label(room, block_key)
        else:
            title = f"{room} - signatures"
            names = ", ".join(duties.get((room, "-")) or [])
            tally = str(room)

        if landscape or row_col or len(rows_df) > TWO_UP_OVER:
            # A list longer than one portrait sheet is pages to keep in order.
            # Two lists side by side on landscape A4 halves them: LHC 110's
            # back blocks (65-70 students) and LHC 105's 80 all come out at two
            # sheets, while a front block of twenty stays on its single page.
            _signature_two_up(cv, rows_df, exam=exam, course=course, when=when,
                              title=title, names=names, tally=tally, serials=serials,
                              landscape=landscape, paper=paper)
            continue

        def head(cont: bool = False):
            _pack_header(cv, f"{course} · {exam}", title + (" (cont.)" if cont else ""),
                         when, "", len(rows_df), tas=names, page_w=pw_sig, page_h=ph_sig)
            y = ph_sig - 56
            cv.setFont("Helvetica-Bold", 8); cv.setFillColorRGB(*MUTE)
            # a serial, not the block and seat: those are on the grid page, and
            # this sheet is read down a list of people. The serial runs on from
            # the previous block and the previous room, so one pile of sheets
            # numbers 1..N over the whole exam.
            cols = ((15, "S. No"), (30, "Roll No"), (62, "Name"), (pw_sig - 70, "Signature"))
            for x, lab in cols:
                cv.drawString(x * mm, y * mm, lab.upper())
            cv.setStrokeColorRGB(*RULE); cv.setLineWidth(0.5)
            cv.line(15 * mm, (y - 2) * mm, (pw_sig - 15) * mm, (y - 2) * mm)
            return y - 2

        y = head()
        for r in rows_df.itertuples():
            y -= 8.4
            if y < 20:
                cv.showPage()
                y = head(cont=True) - 8.4
            cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 8)
            cv.drawString(15 * mm, y * mm, str(serials.get(str(r.Roll), "")))
            cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 8)
            cv.drawString(30 * mm, y * mm, str(r.Roll))
            cv.setFont("Helvetica", 8)
            cv.drawString(62 * mm, y * mm,
                          _fit_text(cv, str(getattr(r, "Name", "")), pw_sig - 15 - 62 - 60, "Helvetica", 8))
            cv.setStrokeColorRGB(0.55, 0.78, 0.88); cv.setLineWidth(0.5)
            cv.rect((pw_sig - 72) * mm, (y - 2.4) * mm, 57 * mm, 7.5 * mm, stroke=1, fill=0)
            cv.setStrokeGray(0.82); cv.setLineWidth(0.3)
            cv.line(15 * mm, (y - 2.6) * mm, (pw_sig - 15) * mm, (y - 2.6) * mm)

        y -= 12
        if y > 20:
            cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 8)
            cv.drawString(15 * mm, y * mm,
                          f"{tally}: {len(rows_df)} students · present ______ · "
                          "absent ______ · scripts collected ______")
        cv.showPage()


def _signature_two_up(cv, rows_df, *, exam, course, when, title, names, tally,
                      serials=None, landscape: bool = False, paper: str = "a4") -> None:
    """The signature list two columns to a landscape sheet.

    Same list, same order, read down the left half and then down the right -
    it is a list to sign, not a page to lay out, and halving the sheet count
    is worth more than the extra width per line.
    """
    is_legal = str(paper).lower() == "legal"
    pw, ph = (LEGAL_W, LEGAL_H) if is_legal else (PAGE_H, PAGE_W)
    top, bottom = ph - 50, 12.0
    row_h = 6.4                                 # still leaves a signable box
    # leave the last line clear of the tally at the foot; 21 per col fits 84 in 2 pages
    per_col = max(21, int((top - bottom - 8) // row_h))
    per_page = per_col * 2
    half = (pw - 30 - 8) / 2                    # 8mm gutter between the columns
    # column offsets within one half: serial, roll, name, signature. The row
    # and column belong on the seating plan, not here - a signature sheet is
    # read down a list of people, and a serial number is what an invigilator
    # counts and what a query later refers to.
    sig_w = 46.0 if is_legal else 39.0
    sig_box_pad = sig_w + 2.0
    offs = (0.0, 10.0, 32.0, half - sig_box_pad)
    total = len(rows_df)
    rows = list(rows_df.itertuples())

    for page_no, start in enumerate(range(0, total, per_page)):
        cv.setPageSize((pw * mm, ph * mm))
        _pack_header(cv, f"{course} · {exam}",
                     title + (" (cont.)" if page_no else ""), when, "", total,
                     tas=names, page_w=pw, page_h=ph)
        chunk = rows[start:start + per_page]
        c_per = (len(chunk) + 1) // 2 if len(chunk) < per_page else per_col
        for col in (0, 1):
            x0 = 15 + col * (half + 8)
            part = chunk[col * c_per:(col + 1) * c_per]
            if not part:
                continue
            cv.setFont("Helvetica-Bold", 7.5); cv.setFillColorRGB(*MUTE)
            for dx, lab in zip(offs, ("S. NO", "ROLL NO", "NAME", "SIGNATURE")):
                cv.drawString((x0 + dx) * mm, top * mm, lab)
            cv.setStrokeColorRGB(*RULE); cv.setLineWidth(0.5)
            cv.line(x0 * mm, (top - 2) * mm, (x0 + half) * mm, (top - 2) * mm)

            y = top - 2
            # the serial runs on across the halves and across the sheets, so
            # "number 47" means one person in the room, not one per column
            n0 = start + col * c_per
            for i, r in enumerate(part, start=n0 + 1):
                y -= row_h
                cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 7.5)
                cv.drawString((x0 + offs[0]) * mm, y * mm,
                              str((serials or {}).get(str(r.Roll), i)))
                cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 7.5)
                cv.drawString((x0 + offs[1]) * mm, y * mm, str(r.Roll))
                cv.setFont("Helvetica", 8.0 if is_legal else 7.5)
                cv.drawString((x0 + offs[2]) * mm, y * mm,
                              _fit_text(cv, str(getattr(r, "Name", "")),
                                        offs[3] - offs[2] - 2, "Helvetica", 8.0 if is_legal else 7.5))
                cv.setStrokeColorRGB(0.55, 0.78, 0.88); cv.setLineWidth(0.5)
                cv.rect((x0 + offs[3]) * mm, (y - 1.8) * mm, sig_w * mm, 5.6 * mm,
                        stroke=1, fill=0)
                cv.setStrokeGray(0.82); cv.setLineWidth(0.3)
                cv.line(x0 * mm, (y - 2.0) * mm, (x0 + half) * mm, (y - 2.0) * mm)

        if start + per_page >= total:
            cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 8)
            cv.drawString(15 * mm, (bottom - 5) * mm,
                          f"{tally}: {total} students · present ______ · "
                          "absent ______ · scripts collected ______")
        cv.showPage()
    if not landscape:
        pw_port, ph_port = (LEGAL_H, LEGAL_W) if is_legal else (PAGE_W, PAGE_H)
        cv.setPageSize((pw_port * mm, ph_port * mm))


def _pack_header(cv, title: str, subtitle: str, when: str, venue: str, n: int,
                 tas: str = "", page_w: float = PAGE_W, page_h: float = PAGE_H) -> None:
    """Four short lines, each on its own row. An earlier version put the
    invigilator names right-aligned on the same line as the counts, which
    collided the moment two TAs were posted to one block."""
    cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica-Bold", 8)
    cv.drawString(15 * mm, (page_h - 18) * mm, title.upper())
    if C.COURSE_INSTRUCTORS:
        cv.setFont("Helvetica", 7.5)
        cv.drawRightString((page_w - 15) * mm, (page_h - 18) * mm, C.instructor_names())
    cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 16)
    cv.drawString(15 * mm, (page_h - 28) * mm, subtitle)

    cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 8.5)
    line = " · ".join(p for p in (when, venue or (f"{n} students" if n else "")) if p)
    cv.drawString(15 * mm, (page_h - 34.5) * mm, line)

    y_rule = page_h - 40
    if tas:
        cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 8.5)
        label = "Invigilators: " if "," in tas else "Invigilator: "
        cv.drawString(15 * mm, (page_h - 41) * mm,
                      label + _fit_text(cv, tas, page_w - 30 - 22, "Helvetica-Bold", 8.5))
        y_rule = page_h - 45
    else:
        cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 8)
        cv.drawString(15 * mm, (page_h - 41) * mm, "Invigilator: ______________________")
        y_rule = page_h - 45

    cv.setStrokeColorRGB(*RULE); cv.setLineWidth(0.8)
    cv.line(15 * mm, y_rule * mm, (page_w - 15) * mm, y_rule * mm)
