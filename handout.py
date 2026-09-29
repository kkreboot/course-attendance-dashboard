"""A minimal student PDF: which hall, which block, and the full name list.

Deliberately plain. No counts, no capacities, no analysis -- a student only
needs their room, their block, and permission to sit anywhere inside it.

Each batch gets a plan page followed by its roster, grouped by block and set in
two columns so the page count stays small.
"""
from __future__ import annotations
from pathlib import Path
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white
import config as C
from config import ROOMS, BLOCK_COLOUR

W, H = 210.0, 297.0
INK = HexColor("#16202B")
MUTE = HexColor("#6B767D")
LINE = HexColor("#B9C0C4")
SLAB = HexColor("#EDEFF0")
GAP = 4.0


def _page(cv, alloc, room, cohort, course, session, ta_lines=None):
    blocks = ROOMS[room]
    rows = [(b, alloc[alloc.Block == b.key].Roll) for b in blocks]
    used = [(b, r) for b, r in rows if len(r)]

    cv.setFillColor(MUTE); cv.setFont("Helvetica-Bold", 7.5)
    cv.drawString(18 * mm, (H - 22) * mm, f"{course.upper()}   ·   {session.upper()}")
    cv.setFillColor(INK); cv.setFont("Helvetica-Bold", 26)
    cv.drawString(18 * mm, (H - 33) * mm, cohort)
    cv.setFont("Helvetica", 13)
    cv.drawString(18 * mm, (H - 42) * mm,
                  room if " " in room else room.replace("LHC", "LHC-"))
    if C.COURSE_INSTRUCTORS:
        cv.setFillColor(MUTE); cv.setFont("Helvetica", 8.5)
        cv.drawRightString((W - 18) * mm, (H - 42) * mm, C.instructor_names())
        cv.setFillColor(INK)
    cv.setStrokeColor(INK); cv.setLineWidth(1.4)
    cv.line(18 * mm, (H - 47) * mm, (W - 18) * mm, (H - 47) * mm)

    cv.setFillColor(INK); cv.setFont("Helvetica-Bold", 10.5)
    cv.drawString(18 * mm, (H - 56) * mm, "Find your roll number below, then sit anywhere inside that block.")

    # ---- block plan -------------------------------------------------------
    # Table height is known, so give the plan whatever is left and let the
    # stage bar follow the last band instead of floating below a gap.
    table_h = 12 + 14 * len(used) + 16
    plan_top, plan_bot = H - 66, 26.0 + table_h
    bands = [(s, [b for b in blocks if b.section == s]) for s in ("rear", "front")]
    bands = [(s, bs) for s, bs in bands if bs]
    avail_h = plan_top - plan_bot - (len(bands) - 1) * 6 - 11
    tot_rows = sum(max(b.rows for b in bs) for _, bs in bands)
    # Keep seats roughly square: the row pitch matches the column pitch, so a
    # tall narrow block on the plan is a tall narrow block in the room.
    scales = [(W - 36) / (sum(b.cols for b in bs) + GAP * (len(bs) - 1)) for _, bs in bands]
    unit = min(min(scales), avail_h / tot_rows)
    plan_h = tot_rows * unit + (len(bands) - 1) * 6 + 11

    y = plan_top - (avail_h + 11 - plan_h) / 2
    for sec, bs in bands:
        bh = max(b.rows for b in bs) * unit
        span = sum(b.cols for b in bs) + GAP * (len(bs) - 1)
        scale = (W - 36) / span
        x = 18.0
        y -= bh
        for b in bs:
            w = b.cols * scale
            on = len(alloc[alloc.Block == b.key]) > 0
            col = HexColor("#" + BLOCK_COLOUR[b.key])
            cv.setFillColor(col if on else SLAB)
            cv.setStrokeColor(col if on else LINE)
            cv.setLineWidth(0.8)
            cv.rect(x * mm, y * mm, w * mm, bh * mm, stroke=1, fill=1)
            cv.setFillColor(white if on else MUTE)
            cv.setFont("Helvetica-Bold", min(26, bh * 1.7))
            cv.drawCentredString((x + w / 2) * mm, (y + bh / 2 - 2.6) * mm, b.key)
            if not on:
                cv.setFont("Helvetica", 6.5)
                cv.drawCentredString((x + w / 2) * mm, (y + bh / 2 - 8) * mm, "not used")
            x += w + GAP * scale
        y -= 6
    y += 6
    cv.setFillColor(INK)
    cv.rect(18 * mm, (y - 11) * mm, (W - 36) * mm, 7 * mm, stroke=0, fill=1)
    cv.setFillColor(white); cv.setFont("Helvetica-Bold", 7.5)
    cv.drawCentredString((W / 2) * mm, (y - 8.6) * mm, "S T A G E   T H I S   S I D E")

    # ---- roll ranges ------------------------------------------------------
    ty = 26.0 + table_h - 12
    cv.setFillColor(MUTE); cv.setFont("Helvetica-Bold", 7)
    cv.drawString(18 * mm, ty * mm, "BLOCK")
    cv.drawString(42 * mm, ty * mm, "ROLL NUMBERS")
    cv.setStrokeColor(LINE); cv.setLineWidth(0.5)
    cv.line(18 * mm, (ty - 3) * mm, (W - 18) * mm, (ty - 3) * mm)
    ty -= 12
    for b, r in used:
        col = HexColor("#" + BLOCK_COLOUR[b.key])
        cv.setFillColor(col)
        cv.rect(18 * mm, (ty - 2.4) * mm, 11 * mm, 11 * mm, stroke=0, fill=1)
        cv.setFillColor(white); cv.setFont("Helvetica-Bold", 13)
        cv.drawCentredString(23.5 * mm, (ty + 1.1) * mm, b.key)
        cv.setFillColor(INK); cv.setFont("Helvetica", 11)
        cv.drawString(42 * mm, (ty + 1.4) * mm, f"{r.min()}   to   {r.max()}")
        cv.setStrokeColor(LINE); cv.setLineWidth(0.4)
        cv.line(18 * mm, (ty - 4.6) * mm, (W - 18) * mm, (ty - 4.6) * mm)
        ty -= 14

    cv.setFillColor(MUTE); cv.setFont("Helvetica-Oblique", 8)
    cv.drawString(18 * mm, 15 * mm,
                  "The full name list for each block follows on the next pages.")
    cv.setFont("Helvetica", 8)
    if ta_lines:
        # Wrap to fit the usable width instead of assuming it fits on one
        # line - the weekday TA roster can run long with two names per day.
        usable = (W - 36) * mm
        rows, cur = [], "Attendance TA - "
        for i, part in enumerate(ta_lines):
            piece = part if cur.endswith("- ") else "   ·   " + part
            if cv.stringWidth(cur + piece, "Helvetica", 8) > usable and not cur.endswith("- "):
                rows.append(cur)
                cur = part
            else:
                cur += piece
        rows.append(cur)
        yy = 10.5
        for row in rows:
            cv.drawString(18 * mm, yy * mm, row)
            yy -= 4.2
    else:
        cv.drawString(18 * mm, 10.5 * mm,
                      "Blocks are fixed for the term. If in doubt, ask a TA before the lecture starts.")
    cv.showPage()


# ---- roster pages ---------------------------------------------------------

COL_TOP, COL_BOT = H - 34, 18.0
ROW_H = 4.9
HEAD_H = 9.5
NCOL = 2
COL_W = (W - 36 - 8) / NCOL


def _roster(cv, alloc, room, cohort):
    """Two-column list, grouped by block, flowing across as many pages as needed."""
    items = []
    for blk in sorted(alloc.Block.unique()):
        sub = alloc[alloc.Block == blk].sort_values("Roll")
        items.append(("head", blk, len(sub)))
        for _, s in sub.iterrows():
            items.append(("row", str(s.Roll), str(s.get("Name", ""))))

    col, y, page_started = 0, COL_TOP, False
    cur_block = None

    def new_page():
        nonlocal col, y, page_started
        if page_started:
            cv.showPage()
        page_started = True
        col, y = 0, COL_TOP
        cv.setFillColor(MUTE); cv.setFont("Helvetica-Bold", 7)
        cv.drawString(18 * mm, (H - 20) * mm,
                      f"{cohort.upper()}   ·   {room if ' ' in room else room.replace('LHC', 'LHC-')}")
        cv.drawRightString((W - 18) * mm, (H - 20) * mm, "SEATING BLOCKS")
        cv.setStrokeColor(LINE); cv.setLineWidth(0.5)
        cv.line(18 * mm, (H - 23) * mm, (W - 18) * mm, (H - 23) * mm)

    def advance(need):
        nonlocal col, y
        if y - need < COL_BOT:
            col += 1
            y = COL_TOP
            if col >= NCOL:
                new_page()

    new_page()
    for it in items:
        if it[0] == "head":
            cur_block = it[1]
            advance(HEAD_H + ROW_H * 2)
            x = 18 + col * (COL_W + 8)
            cv.setFillColor(HexColor("#" + BLOCK_COLOUR[cur_block]))
            cv.rect(x * mm, (y - 6.6) * mm, COL_W * mm, 7.4 * mm, stroke=0, fill=1)
            cv.setFillColor(white); cv.setFont("Helvetica-Bold", 9)
            cv.drawString((x + 2.2) * mm, (y - 4.6) * mm, f"BLOCK {cur_block}")
            cv.setFont("Helvetica", 7.5)
            cv.drawRightString((x + COL_W - 2.2) * mm, (y - 4.6) * mm, f"{it[2]} students")
            y -= HEAD_H
        else:
            advance(ROW_H)
            x = 18 + col * (COL_W + 8)
            if y == COL_TOP and cur_block:      # column break inside a block
                cv.setFillColor(HexColor("#" + BLOCK_COLOUR[cur_block]))
                cv.setFont("Helvetica-Bold", 7)
                cv.drawString(x * mm, (y - 4.2) * mm, f"BLOCK {cur_block} (continued)")
                y -= 7.5
                advance(ROW_H)
                x = 18 + col * (COL_W + 8)
            cv.setFillColor(INK); cv.setFont("Helvetica", 7.4)
            cv.drawString(x * mm, (y - 3.4) * mm, it[1])
            name = it[2]
            while cv.stringWidth(name, "Helvetica", 7.4) > (COL_W - 26) * mm and len(name) > 4:
                name = name[:-2]
            cv.drawString((x + 24) * mm, (y - 3.4) * mm, name)
            cv.setStrokeColor(HexColor("#E4E7E9")); cv.setLineWidth(0.3)
            cv.line(x * mm, (y - 4.9) * mm, (x + COL_W) * mm, (y - 4.9) * mm)
            y -= ROW_H
    cv.showPage()


def build(cohorts, out_pdf: str, *, course=C.COURSE,
          session=C.SESSION, roster=True, tas=None) -> str:
    """cohorts: list of dicts with keys alloc/room/cohort, plus an optional
    `slot` ('Morning' or 'Afternoon') to print that cohort's weekday
    Attendance-TA roster in the footer - this PDF has no single date to look
    a lone TA up against (it's a term-long reference, not a per-session
    sheet), so it shows the whole weekly rotation instead. `tas` is a
    DataFrame from `finder.load_tas`; omit either and the footer falls back
    to the plain "ask a TA" note, unchanged from before this was added."""
    sched = None
    if tas is not None:
        import finder
        sched = finder.ta_schedule(tas)

    cv = canvas.Canvas(out_pdf, pagesize=(W * mm, H * mm))
    cv.setTitle(f"{course} - seating")
    for c in cohorts:
        ta_lines = None
        if sched is not None and c.get("slot"):
            ta_lines = []
            for day in ("Monday", "Wednesday", "Friday"):
                names = sched.get(day, {}).get(c["slot"])
                if names:
                    ta_lines.append(f"{day[:3]} {' & '.join(names)}")
        _page(cv, c["alloc"], c["room"], c["cohort"], course, session, ta_lines=ta_lines)
        if roster:
            _roster(cv, c["alloc"], c["room"], c["cohort"])
    cv.save()
    return out_pdf
