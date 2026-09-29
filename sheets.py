"""Printable blockwise signature sheets, plus the template.json that records
what was printed for each block."""
from __future__ import annotations
import json, math
from pathlib import Path
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import config as C


def _fit(cv, text: str, width_mm: float, font: str, size: float) -> str:
    """Trim `text` to fit `width_mm`, ending in '..' when it had to be cut."""
    if cv.stringWidth(text, font, size) <= width_mm * mm:
        return text
    while text and cv.stringWidth(text + "..", font, size) > width_mm * mm:
        text = text[:-1]
    return text.rstrip() + ".."


# Width of the optional attendance column, taken out of the signature box -
# the box is 73 mm without it, so 20 mm here still leaves more signing room
# than the 53 mm a name gets.
ATT_COL_W = 20.0


def _att_text(rec: dict | None) -> str:
    """`62%  5/9` - the percentage a student is judged on, and the counts it
    came from, so a disputed figure can be argued about concretely."""
    if not rec:
        return "-"
    pct = rec.get("percent")
    held = rec.get("held") or 0
    if pct is None or not held:
        return "-"
    return f"{round(pct):d}%  {rec.get('present', 0)}/{held}"


def build_block_sheet(alloc, block: str, out_pdf: str, *, course: str, room: str,
                      session: str, date_label: str = "", ta_name: str = "",
                      attendance: dict | None = None, att_asof: str = "") -> dict:
    """Write one PDF for one block and return its box template.

    `attendance` ({roll: {percent, present, held}}, from
    `attendance_report.percent_map`) adds a column showing where each student
    stands **up to the last class held** - the sheet they sign on the way into
    the next one. Omit it and the sheet prints exactly as it always has, same
    four columns.
    """
    sub = (alloc[alloc.Block == block]
           .sort_values(["SeatRow", "SeatCol"]).reset_index(drop=True))
    if sub.empty:
        raise ValueError(f"Block {block} has no students.")
    pages = math.ceil(len(sub) / C.ROWS_PER_PAGE)
    cv = canvas.Canvas(out_pdf, pagesize=(C.PAGE_W * mm, C.PAGE_H * mm))
    rows = []

    for p in range(pages):
        chunk = sub.iloc[p * C.ROWS_PER_PAGE:(p + 1) * C.ROWS_PER_PAGE]
        top = C.PAGE_H - C.MARGIN_TOP

        cv.setFont("Helvetica-Bold", 12)
        cv.drawString(C.MARGIN_X * mm, (C.PAGE_H - 31) * mm,
                      f"{course}  |  {room}  |  Block {block}")
        cv.setFont("Helvetica", 8.5)
        cv.drawString(C.MARGIN_X * mm, (C.PAGE_H - 36.5) * mm,
                      f"{session}   ·   page {p+1} of {pages}   ·   "
                      f"seats {chunk.Seat.iloc[0]}–{chunk.Seat.iloc[-1]}")
        cv.drawRightString((C.PAGE_W - C.MARGIN_X) * mm, (C.PAGE_H - 36.5) * mm,
                           f"Date: {date_label or '____________'}")
        if C.COURSE_INSTRUCTORS:
            cv.setFont("Helvetica", 7)
            cv.drawRightString((C.PAGE_W - C.MARGIN_X) * mm, (C.PAGE_H - 31) * mm,
                               C.instructor_names())
        # A TA line uses margin space that's already blank (there's slack between
        # this and the table below) rather than shrinking rows/page, so pages with
        # no TA name print byte-identically to before this was added.
        instr_y = 41.0
        if ta_name:
            cv.setFont("Helvetica-Bold", 8)
            cv.drawString(C.MARGIN_X * mm, (C.PAGE_H - 41) * mm, f"Attendance TA: {ta_name}")
            instr_y = 44.5
        cv.setFont("Helvetica-Oblique", 7)
        cv.drawString(C.MARGIN_X * mm, (C.PAGE_H - instr_y) * mm,
                      "Sign inside your own box only. Do not sign for anyone else.")
        if attendance:
            cv.drawRightString((C.PAGE_W - C.MARGIN_X) * mm, (C.PAGE_H - instr_y) * mm,
                               "Attendance shown is up to "
                               + (f"the class of {att_asof}" if att_asof
                                  else "the last class held")
                               + " - today's is not counted yet.")

        widths = ([C.COL_SEAT, C.COL_ROLL, C.COL_NAME, ATT_COL_W] if attendance
                  else [C.COL_SEAT, C.COL_ROLL, C.COL_NAME])
        sign_w = C.COL_SIGN - (ATT_COL_W if attendance else 0)
        xs = [C.MARGIN_X]
        for w in widths:
            xs.append(xs[-1] + w)
        labels = (("Seat", "Roll No", "Student Name", "Attendance", "Signature")
                  if attendance else ("Seat", "Roll No", "Student Name", "Signature"))
        cv.setFont("Helvetica-Bold", 8)
        cv.setFillGray(0.88)
        cv.rect(C.MARGIN_X * mm, (top - C.ROW_H) * mm,
                (C.PAGE_W - 2 * C.MARGIN_X) * mm, C.ROW_H * mm, stroke=0, fill=1)
        cv.setFillGray(0)
        for x, lab in zip(xs, labels):
            cv.drawString((x + 1.6) * mm, (top - C.ROW_H + 3.0) * mm, lab)

        y = top - C.ROW_H
        for _, s in chunk.iterrows():
            y -= C.ROW_H
            # Seat, roll number and name all in bold: they are what a TA scans
            # down the sheet for, and bold holds up on a photocopy better than
            # 8pt regular does.
            cv.setFont("Helvetica-Bold", 8)
            cv.drawString((xs[0] + 1.6) * mm, (y + 2.9) * mm, str(s.Seat))
            cv.drawString((xs[1] + 1.6) * mm, (y + 2.9) * mm, str(s.Roll))
            # Bold is wider, so trim to the column's actual width rather than a
            # fixed character count, which would let long names run into the
            # signature box.
            cv.drawString((xs[2] + 1.6) * mm, (y + 2.9) * mm,
                          _fit(cv, str(s.Name), C.COL_NAME - 3.2, "Helvetica-Bold", 8))
            rec = attendance.get(str(s.Roll)) if attendance else None
            if attendance:
                # Regular weight, not bold: this is for the student to read
                # once, not for a TA to scan down the column, and bold here
                # competes with the roll number for the eye.
                cv.setFont("Helvetica", 8)
                cv.drawString((xs[3] + 1.6) * mm, (y + 2.9) * mm, _att_text(rec))
            # Signature box in light cyan so the rule never scores as ink.
            cv.setStrokeColorRGB(0.55, 0.78, 0.88)
            cv.setLineWidth(0.6)
            cv.rect(xs[-1] * mm, y * mm, sign_w * mm, C.ROW_H * mm, stroke=1, fill=0)
            cv.setStrokeGray(0.75)
            cv.setLineWidth(0.3)
            cv.line(C.MARGIN_X * mm, y * mm, (C.PAGE_W - C.MARGIN_X) * mm, y * mm)
            row = dict(page=p, seat=str(s.Seat), roll=str(s.Roll), name=str(s.Name))
            if attendance:
                row["attendance"] = _att_text(rec)
            rows.append(row)
        cv.setStrokeGray(0.4)
        cv.setLineWidth(0.6)
        cv.rect(C.MARGIN_X * mm, y * mm, (C.PAGE_W - 2 * C.MARGIN_X) * mm,
                (top - y) * mm, stroke=1, fill=0)
        cv.showPage()
    cv.save()

    return dict(block=block, pdf=Path(out_pdf).name, pages=pages,
                page_mm=[C.PAGE_W, C.PAGE_H],
                course=course, room=room, session=session, rows=rows)


def build_all(alloc, outdir: str, *, course: str, room: str, session: str,
              date_label: str = "", ta_name: str = "",
              attendance: dict | None = None, att_asof: str = "") -> str:
    """One PDF per block plus a single template.json describing the set.

    template.json is what the dashboard discovers cohorts by (`out/<cohort>/
    sheets/template.json`), so it stays even though nothing reads its
    geometry any more."""
    out = Path(outdir)
    out.mkdir(parents=True, exist_ok=True)
    blocks = sorted(alloc.Block.unique())
    # Clear signature PDFs left over from a previous build whose block no
    # longer holds anyone this run - otherwise a stale, un-templated page
    # sits in the folder where it could be printed and handed out by mistake.
    for stale in out.glob("signature_block_*.pdf"):
        if stale.stem[len("signature_block_"):] not in blocks:
            stale.unlink()
    tpl = {"course": course, "room": room, "session": session, "blocks": {}}
    for b in blocks:
        pdf = out / f"signature_block_{b}.pdf"
        tpl["blocks"][b] = build_block_sheet(alloc, b, str(pdf), course=course, room=room,
                                             session=session, date_label=date_label,
                                             ta_name=ta_name, attendance=attendance,
                                             att_asof=att_asof)
    path = out / "template.json"
    path.write_text(json.dumps(tpl, indent=1))
    return str(path)
