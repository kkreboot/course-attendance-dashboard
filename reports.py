"""Deliverables built from an allocation: the seating workbook and the
interactive hall plan."""
from __future__ import annotations
import json
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from config import ROOMS, BLOCK_COLOUR
import config as C
from seating import block_summary

_T = Side(style="thin", color="B0B0B0")
BOX = Border(_T, _T, _T, _T)
HDR = PatternFill("solid", fgColor="16202B")
WEDGE = PatternFill("solid", fgColor="FDEBD2")
YELLOW = PatternFill("solid", fgColor="FFFF00")
ARI = lambda **k: Font(name="Arial", **k)


def build_workbook(alloc, seats, out_xlsx: str, *, room="LHC110", course=C.COURSE,
                   cohort="English Batch", session=C.SESSION,
                   dates=("07th August", "10th August", "12th August", "14th August"),
                   source="") -> str:
    blocks = ROOMS[room]
    summ = block_summary(alloc, seats, room)
    wb = Workbook(); ws = wb.active; ws.title = "Block Summary"
    ws["A1"] = f"{room} · {course} ({cohort}) · Blockwise seating - {len(alloc)} students"
    ws["A1"].font = ARI(bold=True, size=14)
    ws["A2"] = ("Front blocks: every core seat used. Wedge extra seats stay vacant and are "
                "the designated transition seats.")
    ws["A3"] = "Rear blocks: filled by whole rows from the row nearest the front; back rows left empty."
    for r in ("A2", "A3"):
        ws[r].font = ARI(size=9, italic=True, color="555555")

    heads = ["Block", "Position", "Seats/row", "Core", "Wedge", "Capacity", "Seated",
             "Rows used", "Rows total", "Vacant", "First roll", "Last roll"]
    for j, h in enumerate(heads, 1):
        c = ws.cell(5, j, h); c.font = ARI(bold=True, color="FFFFFF"); c.fill = HDR
        c.border = BOX; c.alignment = Alignment("center", "center", wrap_text=True)
    r = 6
    for _, s in summ.iterrows():
        vals = [s.Block, s.Position, s.PerRow, s.Core, s.Wedge, s.Capacity, s.Seated,
                s.RowsUsed, s.RowsTotal, None, s.FirstRoll, s.LastRoll]
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v); c.font = ARI(); c.border = BOX; c.alignment = Alignment("center")
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=BLOCK_COLOUR[s.Block])
        ws.cell(r, 1).font = ARI(bold=True, color="FFFFFF")
        c = ws.cell(r, 10, f"=F{r}-G{r}"); c.font = ARI(bold=True); c.border = BOX
        c.alignment = Alignment("center"); r += 1
    ws.cell(r, 1, "TOTAL").font = ARI(bold=True)
    for j, col in [(4, "D"), (5, "E"), (6, "F"), (7, "G"), (10, "J")]:
        c = ws.cell(r, j, f"=SUM({col}6:{col}{r-1})"); c.font = ARI(bold=True)
        c.border = BOX; c.alignment = Alignment("center")
    for j in (1, 2, 3, 8, 9):
        ws.cell(r, j).border = BOX

    wedge = seats[seats.Kind.eq("extra")]
    ranges = ", ".join(f"{b}: {g.Seat.min()}–{g.Seat.max()} ({len(g)})"
                       for b, g in wedge.groupby("Block"))
    notes = [f"Transition seats, allot in this order - {ranges}.",
             "Assumed geometry: " + "; ".join(f"{b.key} {b.cols} wide × {b.rows} rows" for b in blocks) + ".",
             "TA duty is not assigned here. See the separate TA duty workbook.",
             f"Source: {source}" if source else ""]
    for k, n in enumerate(x for x in notes if x):
        ws.cell(r + 2 + k, 1, n).font = ARI(size=9, italic=True, color="777777")
    for j, w in enumerate([8, 15, 10, 8, 8, 10, 9, 10, 10, 9, 13, 13], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A6"

    cols = [c for c in ["Roll", "Name", "Program", "Group", "Medium"] if c in alloc]
    ms = wb.create_sheet("Master (roll order)")
    heads = ["Sr"] + [{"Roll": "Roll No", "Name": "Student Name", "Program": "Program",
                       "Group": "Group", "Medium": "Medium"}[c] for c in cols] + \
            ["Block", "Seat", "Row", "Col"]
    for j, h in enumerate(heads, 1):
        c = ms.cell(1, j, h); c.font = ARI(bold=True, color="FFFFFF"); c.fill = HDR
        c.border = BOX; c.alignment = Alignment("center")
    for i, (_, s) in enumerate(alloc.sort_values("Roll").iterrows(), start=2):
        vals = [i - 1] + [s[c] for c in cols] + [s.Block, s.Seat, s.SeatRow, s.SeatCol]
        for j, v in enumerate(vals, 1):
            c = ms.cell(i, j, v); c.font = ARI(size=10); c.border = BOX
            if j == 1 or j > len(cols) + 1 or heads[j-1] in ("Group", "Medium"):
                c.alignment = Alignment("center")
        bj = len(cols) + 2
        ms.cell(i, bj).fill = PatternFill("solid", fgColor=BLOCK_COLOUR[s.Block])
        ms.cell(i, bj).font = ARI(bold=True, color="FFFFFF", size=10)
    widths = [5] + [{"Roll": 13, "Name": 32, "Program": 36, "Group": 7, "Medium": 9}[c]
                    for c in cols] + [7, 9, 6, 6]
    for j, w in enumerate(widths, 1):
        ms.column_dimensions[get_column_letter(j)].width = w
    ms.freeze_panes = "A2"
    ms.auto_filter.ref = f"A1:{get_column_letter(len(heads))}{len(alloc)+1}"

    tr = wb.create_sheet("Transition seats")
    tr["A1"] = "Vacant wedge seats - held for medium transfers and late admissions"
    tr["A1"].font = ARI(bold=True, size=13)
    tr["A2"] = "Allot in the order listed, then add the student to the Master sheet."
    tr["A2"].font = ARI(size=9, italic=True, color="555555")
    for j, h in enumerate(["Seat", "Block", "Row", "Allotted to (Roll No)",
                           "Student Name", "Date"], 1):
        c = tr.cell(4, j, h); c.font = ARI(bold=True, color="FFFFFF"); c.fill = HDR
        c.border = BOX; c.alignment = Alignment("center")
    i = 5
    for _, x in wedge.sort_values(["Block", "Row", "Col"]).iterrows():
        for j, v in enumerate([x.Seat, x.Block, int(x.Row), None, None, None], 1):
            c = tr.cell(i, j, v); c.font = ARI(size=10); c.border = BOX
            if j <= 3:
                c.alignment = Alignment("center")
            if j >= 4:
                c.fill = YELLOW; c.font = ARI(size=10, color="0000FF")
        tr.cell(i, 1).fill = WEDGE; i += 1
    tr.cell(i + 1, 1, f"Yellow cells are for you to fill in. {len(wedge)} seats total.").font = \
        ARI(size=9, italic=True)
    for j, w in enumerate([10, 8, 7, 22, 30, 12], 1):
        tr.column_dimensions[get_column_letter(j)].width = w
    tr.freeze_panes = "A5"

    for b in ROOMS[room]:
        sub = alloc[alloc.Block == b.key].sort_values(["SeatRow", "SeatCol"])
        if sub.empty:
            continue
        s = wb.create_sheet(f"Block {b.key}")
        s["A1"] = f"{room}  |  Block {b.key} ({b.side})  |  {course} - {cohort}"
        s["A1"].font = ARI(bold=True, size=13)
        s["A2"] = (f"{len(sub)} students · rows 1–{int(sub.SeatRow.max())} of {b.rows} · "
                   "Row 1 nearest the stage · sign only against your own roll number")
        s["A2"].font = ARI(size=9, italic=True, color="555555")
        head = ["Seat", "Row", "Roll No", "Student Name", "Group"] + list(dates)
        for j, x in enumerate(head, 1):
            c = s.cell(4, j, x); c.font = ARI(bold=True, color="FFFFFF"); c.fill = HDR
            c.border = BOX; c.alignment = Alignment("center", "center", wrap_text=True)
        i, prev = 5, None
        for _, st in sub.iterrows():
            for j, v in enumerate([st.Seat, st.SeatRow, st.Roll, st.get("Name", ""),
                                   st.get("Group", "")], 1):
                c = s.cell(i, j, v); c.font = ARI(size=10); c.border = BOX
                if j in (1, 2, 5):
                    c.alignment = Alignment("center")
            if prev is not None and st.SeatRow != prev:
                for j in range(1, 6 + len(dates)):
                    s.cell(i, j).border = Border(_T, _T, Side(style="medium", color="16202B"), _T)
            for j in range(6, 6 + len(dates)):
                s.cell(i, j).border = BOX
            s.row_dimensions[i].height = 22
            prev = st.SeatRow; i += 1
        for j, w in enumerate([10, 6, 13, 34, 7], 1):
            s.column_dimensions[get_column_letter(j)].width = w
        for j in range(6, 6 + len(dates)):
            s.column_dimensions[get_column_letter(j)].width = 15
        s.freeze_panes = "A5"; s.print_title_rows = "4:4"
        s.page_setup.orientation = "landscape"; s.page_setup.fitToWidth = 1
        s.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(out_xlsx)
    return out_xlsx


def build_html(alloc, seats, out_html: str, *, room="LHC110", course=C.COURSE,
               cohort="English batch", session=C.SESSION) -> str:
    """Interactive hall plan: click or search a roll number to spotlight a seat."""
    blocks = ROOMS[room]
    occ = {r.Seat: r for _, r in alloc.iterrows()}
    payload = {"blocks": {}, "meta": {
        "n": len(alloc), "room": room, "course": course, "cohort": cohort,
        "session": session,
        "cap": int(sum(b.total for b in blocks)),
        "wedge": int(sum(b.extra for b in blocks)),
        "en": int((alloc.get("Medium", pd.Series(dtype=str)) == "English").sum()),
        "hi": int((alloc.get("Medium", pd.Series(dtype=str)) == "Hindi").sum())}}
    for b in blocks:
        rows = []
        for _, r in seats[seats.Block == b.key].iterrows():
            if r.Seat in occ:
                s = occ[r.Seat]
                rows.append({"c": r.Seat, "r": int(r.Row), "x": int(r.Col), "k": r.Kind, "v": 0,
                             "roll": s.Roll, "n": s.get("Name", ""), "p": s.get("Program", ""),
                             "g": s.get("Group", ""), "m": s.get("Medium", "")})
            else:
                rows.append({"c": r.Seat, "r": int(r.Row), "x": int(r.Col), "k": r.Kind, "v": 1})
        sub = alloc[alloc.Block == b.key]
        payload["blocks"][b.key] = {
            "side": b.side, "total": b.total, "rows": b.rows, "extra": b.extra,
            "used": int(sub.SeatRow.max()) if len(sub) else 0, "occ": len(sub),
            "seats": sorted(rows, key=lambda s: (s["r"], s["x"]))}
    tpl_path = Path(__file__).parent / "plan_template.html"
    if not tpl_path.exists():
        html = f"<html><body><h1>{room} - {course} ({cohort})</h1><p>Data allocated successfully for {len(alloc)} students.</p></body></html>"
        Path(out_html).write_text(html, encoding="utf-8")
        return out_html
    tpl = tpl_path.read_text(encoding="utf-8")
    order = [b.key for b in blocks]
    html = (tpl.replace("__DATA__", json.dumps(payload, ensure_ascii=False, separators=(",", ":")))
               .replace("__ORDER__", json.dumps(order))
               .replace("__FRONT__", json.dumps([b.key for b in blocks if b.section == "front"]))
               .replace("__REAR__", json.dumps([b.key for b in blocks if b.section == "rear"]))
               .replace("__TITLE__", f"{room} · {course}")
               .replace("__COURSE__", course).replace("__COHORT__", cohort)
               .replace("__SESSION__", session).replace("__ROOM__", room))
    Path(out_html).write_text(html, encoding="utf-8")
    return out_html
