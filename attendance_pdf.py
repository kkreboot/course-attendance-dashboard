#!/usr/bin/env python3
"""The attendance report as a PDF, built from the register rather than typed.

The course has been producing this by hand once a fortnight: where the class
stands, which students are below the requirement, which of them can still get
back to it, and how each class turned out. Every number in it already exists in
`attendance_report` and `mailer`, so the report is a drawing job, not a second
calculation - and that matters, because a report whose figures are worked out
separately will eventually disagree with the mail a student receives.

Four sections, each of which earns its page:

  * **where the class stands** - four figures and a batch comparison, the
    numbers someone asks for first;
  * **class by class** - turnout per session, with the ones under the
    benchmark marked, which is how a pattern (a bad slot, a festival week)
    becomes visible;
  * **who is below the requirement** - the list, lowest first, each row saying
    how many consecutive classes would get that student back and the best they
    can still reach, colour-coded by whether that is possible at all;
  * **method** - the formula, what an excused mark does, where the numbers
    came from and who is a special case. A report that states its own method
    can be checked; one that doesn't has to be trusted.
"""
from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

import attendance_report as attrep
import config as C
import mailer

PAGE_W, PAGE_H = 210.0, 297.0
MARGIN = 15.0

INK, MUTE, RULE = (0.13, 0.13, 0.11), (0.42, 0.40, 0.35), (0.80, 0.78, 0.72)
# Amber can still recover, purple cannot. The same two colours the course has
# been using on the sheet it reads these from.
AMBER_FILL, AMBER_INK = (0.99, 0.94, 0.82), (0.55, 0.36, 0.05)
PURPLE_FILL, PURPLE_INK = (0.93, 0.90, 0.97), (0.38, 0.24, 0.60)
TILE_FILL = (0.95, 0.96, 0.95)

DEFAULT_BENCHMARK = 85.0


def _fmt_pct(value, dp: int = 1) -> str:
    if value is None:
        return "-"
    out = f"{float(value):.{dp}f}"
    return out[:-2] if dp and out.endswith("." + "0" * dp) else out


def _fit(cv, text: str, width_mm: float, font: str, size: float) -> str:
    text = str(text)
    if cv.stringWidth(text, font, size) <= width_mm * mm:
        return text
    while text and cv.stringWidth(text + "…", font, size) > width_mm * mm:
        text = text[:-1]
    return text + "…"


def _wrap(cv, text: str, width_mm: float, font: str, size: float) -> list[str]:
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


def _date(value, fmt: str = "%d/%m/%Y") -> str:
    if isinstance(value, (date, datetime)):
        return value.strftime(fmt)
    return str(value or "")


# ───────────────────────────── the figures ──────────────────────────────
def report_data(path: str, *, sheet: str = "", threshold: float | None = None,
                total_classes: int = 0, benchmark: float | None = None,
                excused: str = attrep.EXCUSED_MODE) -> dict:
    """Everything the report prints, read once from the register.

    Settings come from the workbook's own Dashboard sheet where it has one, so
    the report answers the question the course set rather than one this module
    invented; anything passed in overrides that.
    """
    blocks = attrep.dashboard_blocks(path)
    settings = blocks.get("settings", {})
    threshold = float(threshold if threshold is not None
                      else settings.get("threshold (%)") or 75.0)
    benchmark = float(benchmark if benchmark is not None
                      else settings.get("class benchmark (%)") or DEFAULT_BENCHMARK)

    sheet = sheet or (attrep.combined_sheets(path) or attrep.cohort_sheets(path))[0]
    students, sessions = attrep.load(path, sheet, excused=excused)

    # per-batch turnout per session: the combined sheet cannot give it, since a
    # session column there is both batches at once
    per_batch: dict[str, dict] = {}
    for name in attrep.cohort_sheets(path):
        _, sess = attrep.load(path, name, excused=excused)
        per_batch[name] = {r.Session: r.Percent for r in sess.itertuples()}

    marked = sessions[(sessions.Present + sessions.Absent + sessions.Excused) > 0]
    held = len(marked)
    total = int(total_classes or float(settings.get("total classes (course)") or 0)
                or len(sessions))
    remaining = max(0, total - held)
    asof = _date(marked.Date.max()) if held else ""
    last_label = str(marked.Session.iloc[-1]) if held else ""

    seen = students[students.Held > 0]
    below = seen[seen.Percent < threshold].sort_values("Percent").copy()
    below["Needed"] = [mailer.classes_to_recover(int(r.Present), int(r.Absent), threshold)
                       for r in below.itertuples()]
    below["Best"] = [mailer.best_possible_percent(int(r.Present), int(r.Absent), remaining)
                     for r in below.itertuples()]
    below["Recoverable"] = [mailer.can_still_reach(int(r.Present), int(r.Absent),
                                                   remaining, threshold)
                            for r in below.itertuples()]

    # batch order follows the workbook's own sheet order, not the order the
    # first student happens to appear in on the combined list
    seen_batches = {str(b).strip() for b in students.Batch if str(b).strip()}
    batches = [b for b in attrep.cohort_sheets(path) if b in seen_batches] or \
        [b for b in dict.fromkeys(students.Batch) if str(b).strip()]
    rows = []
    for label, part in [("Both batches", seen)] + \
            [(b, seen[seen.Batch == b]) for b in batches]:
        if not len(part):
            continue
        part_below = part[part.Percent < threshold]
        rows.append({
            "batch": label, "students": len(part),
            "below": len(part_below),
            "share": len(part_below) / len(part) * 100 if len(part) else 0.0,
            "critical": int((part.Percent < 50).sum()),
            "cannot": int((~part_below.index.isin(below[below.Recoverable].index)).sum())
            if len(part_below) else 0,
            "average": float(part.Percent.mean()),
        })

    classes = []
    for r in marked.itertuples():
        classes.append({
            "session": str(r.Session), "date": _date(r.Date, "%d/%m"),
            "present": int(r.Present), "absent": int(r.Absent),
            "excused": int(getattr(r, "Excused", 0)),
            "percent": float(r.Percent) if r.Percent is not None else None,
            "by_batch": {b: per_batch.get(b, {}).get(r.Session) for b in batches},
        })

    # a dated column with nothing in it is a class that did not happen, or one
    # nobody has entered yet. Either way it is not "held", and a report that
    # counts 22 classes while labelling the last one C23 has to say why.
    gaps = [str(r.Session) for r in sessions.itertuples()
            if str(r.Session) not in set(marked.Session)
            and r.Date is not None and str(r.Session) <= str(last_label)]

    return {
        "sheet": sheet, "threshold": threshold, "benchmark": benchmark,
        "gaps": gaps,
        "students": students, "seen": seen, "below": below, "sessions": sessions,
        "held": held, "total": total, "remaining": remaining,
        "asof": asof, "last_label": last_label, "batches": batches,
        "summary": rows, "classes": classes,
        "average": float(seen.Percent.mean()) if len(seen) else None,
        "critical": int((seen.Percent < 50).sum()) if len(seen) else 0,
        "cannot_reach": int((~below.Recoverable).sum()) if len(below) else 0,
        "excused_marks": int(students.Excused.sum()),
        "excused_students": int((students.Excused > 0).sum()),
        "partial": students[(students.NotApplicable > 0)],
        "read_on": date.today().strftime("%d/%m/%Y"),
        "excused_mode": excused,
    }


def observations(d: dict) -> list[str]:
    """The paragraph a reader would otherwise have to work out from the tables.

    Only what the numbers support: no advice, no cause, and nothing that would
    quietly become false next fortnight.
    """
    out = []
    thr, below = d["threshold"], d["below"]
    seen = d["seen"]
    if len(below):
        by_batch = below.Batch.value_counts().to_dict()
        split = ", ".join(f"{n} in {b}" for b, n in by_batch.items())
        out.append(f"{len(below)} students ({_fmt_pct(len(below) / max(len(seen), 1) * 100)}%) "
                   f"are below {_fmt_pct(thr)}%: {split}.")
        crit = seen[seen.Percent < 50]
        if len(crit):
            where = crit.Batch.value_counts().to_dict()
            out.append(f"{len(crit)} are below 50%"
                       + (f", all in {list(where)[0]}" if len(where) == 1 else
                          ": " + ", ".join(f"{n} in {b}" for b, n in where.items())) + ".")
        lost = below[~below.Recoverable]
        if len(lost):
            best = lost.Best.dropna()
            out.append(f"{len(lost)} cannot reach {_fmt_pct(thr)}% even by attending all "
                       f"{d['remaining']} remaining classes"
                       + (f" (best case {_fmt_pct(best.min())}% to {_fmt_pct(best.max())}%)"
                          if len(best) else "") + ".")
        rec = below[below.Recoverable]
        if len(rec):
            out.append(f"The other {len(rec)} can still recover; they need between "
                       f"{int(rec.Needed.min())} and {int(rec.Needed.max())} consecutive "
                       "classes.")
    else:
        out.append(f"Nobody is below {_fmt_pct(thr)}%.")

    if len(d["batches"]) > 1:
        avgs = {b: seen[seen.Batch == b].Percent.mean() for b in d["batches"]
                if len(seen[seen.Batch == b])}
        if len(avgs) > 1:
            ranked = sorted(avgs.items(), key=lambda kv: -kv[1])
            out.append(" vs ".join(f"{b} averages {_fmt_pct(v)}%" for b, v in ranked) + ".")

    low = [c for c in d["classes"] if c["percent"] is not None]
    if low:
        worst = sorted(low, key=lambda c: c["percent"])[:3]
        out.append("Lowest turnout: "
                   + ", ".join(f"{c['session']} ({c['date']}) {_fmt_pct(c['percent'])}%"
                               for c in worst) + ".")
        under = [c for c in low if c["percent"] < d["benchmark"]]
        if under:
            out.append(f"{len(under)} of {len(low)} classes were under the "
                       f"{_fmt_pct(d['benchmark'])}% class benchmark: "
                       + ", ".join(c["session"] for c in under) + ".")
    return out


# ────────────────────────────── the drawing ─────────────────────────────
class _Sheet:
    """One canvas, with the running header and page number this report uses."""

    def __init__(self, cv, title: str):
        self.cv, self.title, self.page = cv, title, 0
        self.y = 0.0
        self.new_page()

    def new_page(self) -> None:
        cv = self.cv
        if self.page:
            cv.showPage()
        self.page += 1
        cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 7.5)
        cv.drawString(MARGIN * mm, (PAGE_H - 12) * mm, self.title)
        cv.drawRightString((PAGE_W - MARGIN) * mm, (PAGE_H - 12) * mm, f"Page {self.page}")
        cv.setStrokeColorRGB(*RULE); cv.setLineWidth(0.4)
        cv.line(MARGIN * mm, (PAGE_H - 14) * mm, (PAGE_W - MARGIN) * mm, (PAGE_H - 14) * mm)
        self.y = PAGE_H - 24

    def room(self, need: float) -> None:
        if self.y - need < 18:
            self.new_page()

    def heading(self, text: str, size: float = 13) -> None:
        self.room(14)
        self.cv.setFillColorRGB(*INK); self.cv.setFont("Helvetica-Bold", size)
        self.cv.drawString(MARGIN * mm, self.y * mm, text)
        self.y -= size * 0.45 + 3


def _tiles(sh: _Sheet, tiles: list[tuple[str, str]]) -> None:
    """The four figures someone asks for first, big enough to read across a
    desk."""
    cv = sh.cv
    usable = PAGE_W - 2 * MARGIN
    gap, n = 4.0, max(len(tiles), 1)
    w = (usable - gap * (n - 1)) / n
    h = 22.0
    sh.room(h + 4)
    for i, (value, label) in enumerate(tiles):
        x = MARGIN + i * (w + gap)
        cv.setFillColorRGB(*TILE_FILL); cv.setStrokeColorRGB(*RULE); cv.setLineWidth(0.4)
        cv.rect(x * mm, (sh.y - h) * mm, w * mm, h * mm, stroke=1, fill=1)
        cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 17)
        cv.drawString((x + 3) * mm, (sh.y - 10) * mm, str(value))
        cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 7.5)
        for j, line in enumerate(_wrap(cv, label, w - 6, "Helvetica", 7.5)[:2]):
            cv.drawString((x + 3) * mm, (sh.y - 15 - j * 3.6) * mm, line)
    sh.y -= h + 6


def _table(sh: _Sheet, headers: list[tuple[float, str, str]], rows: list[list],
           row_h: float = 5.6, size: float = 7.5, fills=None) -> None:
    """`headers` is (x offset in mm, label, align) with align in 'l'/'r'.
    `fills` gives an optional (fill, ink) per row."""
    cv = sh.cv

    def head():
        sh.room(row_h * 2)
        cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica-Bold", size - 0.7)
        for x, label, align in headers:
            if align == "r":
                cv.drawRightString((MARGIN + x) * mm, sh.y * mm, label.upper())
            else:
                cv.drawString((MARGIN + x) * mm, sh.y * mm, label.upper())
        sh.y -= 1.8
        cv.setStrokeColorRGB(*RULE); cv.setLineWidth(0.4)
        cv.line(MARGIN * mm, sh.y * mm, (PAGE_W - MARGIN) * mm, sh.y * mm)
        sh.y -= row_h

    head()
    for i, row in enumerate(rows):
        if sh.y < 20:
            sh.new_page()
            head()
        fill, ink = (fills[i] if fills else (None, INK))
        if fill:
            cv.setFillColorRGB(*fill)
            cv.rect(MARGIN * mm, (sh.y - 1.6) * mm, (PAGE_W - 2 * MARGIN) * mm,
                    (row_h - 0.6) * mm, stroke=0, fill=1)
        for (x, _, align), value in zip(headers, row):
            cv.setFillColorRGB(*ink); cv.setFont("Helvetica", size)
            text = str(value)
            if align == "r":
                cv.drawRightString((MARGIN + x) * mm, sh.y * mm, text)
            else:
                width = (headers[headers.index((x, _, align)) + 1][0] - x - 2
                         if headers.index((x, _, align)) + 1 < len(headers)
                         else PAGE_W - 2 * MARGIN - x)
                cv.drawString((MARGIN + x) * mm, sh.y * mm,
                              _fit(cv, text, width, "Helvetica", size))
        sh.y -= row_h
    sh.y -= 2


def _bullets(sh: _Sheet, lines: list[str], size: float = 8.5) -> None:
    cv = sh.cv
    for line in lines:
        wrapped = _wrap(cv, line, PAGE_W - 2 * MARGIN - 4, "Helvetica", size)
        sh.room(len(wrapped) * 4.2 + 1)
        for j, part in enumerate(wrapped):
            cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", size)
            cv.drawString((MARGIN + (0 if j else 0)) * mm, sh.y * mm,
                          ("· " if j == 0 else "  ") + part)
            sh.y -= 4.2
        sh.y -= 1.2


def build_report(path: str, out_pdf: str, *, sheet: str = "", threshold: float | None = None,
                 total_classes: int = 0, benchmark: float | None = None,
                 course: str = C.COURSE, session: str = C.SESSION,
                 class_wise: bool = True, method: bool = True,
                 excused: str = attrep.EXCUSED_MODE,
                 data: dict | None = None) -> str:
    """Write the report. Returns the path written."""
    d = data or report_data(path, sheet=sheet, threshold=threshold,
                            total_classes=total_classes, benchmark=benchmark,
                            excused=excused)
    thr = d["threshold"]
    title = (f"{course} · {session} · attendance summary"
             + (f" (data up to {d['asof']})" if d["asof"] else ""))

    cv = canvas.Canvas(out_pdf, pagesize=(PAGE_W * mm, PAGE_H * mm))
    cv.setTitle(f"{course} - attendance summary")
    sh = _Sheet(cv, title)

    # ── where the class stands ───────────────────────────────────────────
    cv.setFillColorRGB(*INK); cv.setFont("Helvetica-Bold", 20)
    cv.drawString(MARGIN * mm, sh.y * mm, f"{course.split()[0]} attendance summary")
    sh.y -= 7
    cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 9)
    sub = (f"{session} · {' + '.join(d['batches'])} · "
           + (f"data up to {d['last_label']} ({d['asof']}) · " if d["asof"] else "")
           + f"{d['held']} of {d['total']} classes held, {d['remaining']} remaining")
    for line in _wrap(cv, sub, PAGE_W - 2 * MARGIN, "Helvetica", 9):
        cv.drawString(MARGIN * mm, sh.y * mm, line)
        sh.y -= 4.4
    if C.COURSE_INSTRUCTORS:
        cv.drawString(MARGIN * mm, sh.y * mm, f"Course instructors: {C.instructor_names()}")
        sh.y -= 4.4
    sh.y -= 3

    _tiles(sh, [
        (str(len(d["students"])), "Students enrolled"),
        (f"{_fmt_pct(d['average'])}%", "Average attendance"),
        (str(len(d["below"])), f"Below {_fmt_pct(thr)}% now"),
        (str(d["cannot_reach"]), f"Cannot reach {_fmt_pct(thr)}% by course end"),
    ])

    sh.heading("Batch comparison")
    _table(sh, [(0, "Batch", "l"), (52, "Students", "r"), (76, f"Below {_fmt_pct(thr)}%", "r"),
                (98, "% of batch", "r"), (122, "Below 50%", "r"),
                (150, f"Cannot reach {_fmt_pct(thr)}%", "r"),
                (180, "Average", "r")],
           [[r["batch"], r["students"], r["below"], f"{_fmt_pct(r['share'])}%",
             r["critical"], r["cannot"], f"{_fmt_pct(r['average'], 2)}%"]
            for r in d["summary"]], row_h=6.0, size=8)

    bands = [("Below 50%", 0, 50), (f"50 to {_fmt_pct(thr)}%", 50, thr),
             (f"{_fmt_pct(thr)} to 90%", thr, 90), ("90% and above", 90, 101)]
    seen = d["seen"]
    sh.heading("Attendance bands")
    _table(sh, [(0, "Band", "l"), (60, "Students", "r"), (82, "Share", "r")],
           [[name, int(((seen.Percent >= lo) & (seen.Percent < hi)).sum()),
             f"{_fmt_pct(((seen.Percent >= lo) & (seen.Percent < hi)).sum() / max(len(seen), 1) * 100)}%"]
            for name, lo, hi in bands], row_h=5.6, size=8)

    sh.heading("Key observations")
    _bullets(sh, observations(d))

    # ── class by class ───────────────────────────────────────────────────
    if class_wise and d["classes"]:
        sh.new_page()
        sh.heading("Class-wise attendance")
        cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 8)
        cv.drawString(MARGIN * mm, sh.y * mm,
                      f"Turnout per class held. A class under the "
                      f"{_fmt_pct(d['benchmark'])}% benchmark is marked.")
        sh.y -= 6
        headers = [(0, "Class", "l"), (20, "Date", "l"), (44, "Present", "r"),
                   (62, "Absent", "r"), (80, "Excused", "r"), (100, "All", "r")]
        for i, b in enumerate(d["batches"][:2]):
            headers.append((120 + i * 24, b, "r"))
        headers.append((PAGE_W - 2 * MARGIN, "Note", "r"))
        rows, fills = [], []
        for c in d["classes"]:
            under = c["percent"] is not None and c["percent"] < d["benchmark"]
            row = [c["session"], c["date"], c["present"], c["absent"], c["excused"],
                   f"{_fmt_pct(c['percent'])}%"]
            for b in d["batches"][:2]:
                row.append(f"{_fmt_pct(c['by_batch'].get(b))}%")
            row.append("under benchmark" if under else "")
            rows.append(row)
            fills.append((AMBER_FILL, AMBER_INK) if under else (None, INK))
        _table(sh, headers, rows, row_h=5.4, size=7.5, fills=fills)

    # ── who is below the requirement ─────────────────────────────────────
    sh.new_page()
    sh.heading(f"Students below {_fmt_pct(thr)}% ({len(d['below'])})")
    cv.setFillColorRGB(*MUTE); cv.setFont("Helvetica", 8)
    for line in _wrap(cv, "Lowest first. Classes needed is the run of consecutive classes "
                          f"that brings a student back to {_fmt_pct(thr)}%; best possible is "
                          f"where they finish if they attend all {d['remaining']} remaining "
                          "classes. Purple rows cannot reach the requirement, amber rows can.",
                      PAGE_W - 2 * MARGIN, "Helvetica", 8):
        cv.drawString(MARGIN * mm, sh.y * mm, line)
        sh.y -= 4
    sh.y -= 2

    headers = [(0, "#", "l"), (8, "Roll no.", "l"), (36, "Name", "l"), (94, "Batch", "l"),
               (122, "Att. %", "r"), (134, "Y", "r"), (144, "N", "r"), (154, "E", "r"),
               (168, "Needed", "r"), (PAGE_W - 2 * MARGIN, "Best", "r")]
    rows, fills = [], []
    for i, r in enumerate(d["below"].itertuples(), start=1):
        rows.append([i, r.Roll, r.Name, r.Batch, _fmt_pct(r.Percent), int(r.Present),
                     int(r.Absent), int(r.Excused), int(r.Needed), f"{_fmt_pct(r.Best)}%"])
        fills.append((AMBER_FILL, AMBER_INK) if r.Recoverable else (PURPLE_FILL, PURPLE_INK))
    if rows:
        _table(sh, headers, rows, row_h=5.6, size=7.5, fills=fills)
    else:
        cv.setFillColorRGB(*INK); cv.setFont("Helvetica", 9)
        cv.drawString(MARGIN * mm, sh.y * mm,
                      f"Nobody is below {_fmt_pct(thr)}%.")
        sh.y -= 6

    # ── method ───────────────────────────────────────────────────────────
    if method:
        # let it start wherever it fits rather than reserving a whole page: an
        # almost empty last sheet is paper nobody wanted
        sh.room(24)
        sh.heading("Method and notes")
        excl = ("left out of the percentage entirely" if d["excused_mode"] == "exclude"
                else "counted as present")
        notes = [
            f"Source: the {d['sheet']} sheet of the Institute's Class Attendance workbook, "
            f"read on {d['read_on']}.",
            f"Attendance % = Y / (Y + N) x 100, recomputed here rather than read from the "
            "sheet's cached formula cells.",
            f"An excused mark (E: medical or another accepted reason) is never an absence "
            f"and is {excl}. {d['excused_marks']} E marks across {d['excused_students']} "
            "students.",
            f"Course length is taken as {d['total']} classes; {d['held']} have been held, "
            f"so {d['remaining']} remain. Best possible assumes every remaining class is "
            "attended.",
            f"Classes needed is the smallest k for which (Y + k) / (Y + N + k) >= "
            f"{_fmt_pct(thr)}%.",
        ]
        if d.get("gaps"):
            notes.append(f"{', '.join(d['gaps'])} " +
                         ("is a dated column with no marks" if len(d["gaps"]) == 1
                          else "are dated columns with no marks") +
                         ", so " + ("it is" if len(d["gaps"]) == 1 else "they are") +
                         " not counted as held.")
        for r in d["partial"].itertuples():
            notes.append(f"{r.Name} ({r.Roll}) has {int(r.NotApplicable)} classes marked N/A "
                         "(before joining the roll), so the calculation uses the "
                         f"{int(r.Held)} classes that apply to them.")
        _bullets(sh, notes, size=8)

    cv.showPage()
    cv.save()
    return out_pdf


if __name__ == "__main__":                              # pragma: no cover - manual use
    import sys
    wb = sys.argv[1] if len(sys.argv) > 1 else C.ATTENDANCE_WORKBOOK
    out = sys.argv[2] if len(sys.argv) > 2 else "attendance_report.pdf"
    print(build_report(wb, out))
