"""A single hostable page: a student types their roll number and gets their
room and block, with the block highlighted on the hall plan. A TA can type
their roll number instead and get their duty assignment.

Roll numbers are stored as hashes, never in plain text, and no names or
programmes are embedded. Lookup is exact-match only, so the page cannot be
browsed as a student (or TA) directory. Roll numbers are guessable by
pattern, so this is protection against casual scraping rather than a
security boundary.
"""
from __future__ import annotations
import json
from pathlib import Path

import pandas as pd

from config import ROOMS, BLOCK_COLOUR
import config as C

M32 = 0xFFFFFFFF


def _fnv(text: str, basis: int) -> int:
    h = basis
    for b in text.encode("utf-8"):
        h ^= b
        h = (h * 16777619) & M32
    return h


def roll_key(roll: str) -> str:
    """Two FNV-1a passes with different bases, concatenated. Mirrored in JS."""
    t = roll.strip().upper()
    return f"{_fnv(t, 2166136261):08x}{_fnv(t, 40389):08x}"


def load_tas(path: str, sheet=0) -> pd.DataFrame:
    """Reads the Institute's own TA duty assignment format: a title row, a
    blank row, then a header row (whose 5th column, the day, is printed with
    no header text of its own) followed by one row per TA. Returns
    Roll/Name/Duty/Day/Notes - Day and Notes are blank where the sheet
    leaves them blank (e.g. the two "Head" rows have a description in Notes
    and no Day; Attendance-duty rows have a Day and an "Morning"/"Afternoon"
    slot in Notes instead)."""
    raw = pd.read_excel(path, sheet_name=sheet, header=None)
    header_row = next(i for i in range(min(6, len(raw))) if str(raw.iloc[i, 1]).strip().lower() in
                      ("ta roll", "roll", "roll no", "roll no."))
    data = raw.iloc[header_row + 1:]
    df = pd.DataFrame({
        "Roll": data[1].astype(str).str.strip(),
        "Name": data[2].astype(str).str.strip(),
        "Duty": data[3].astype(str).str.strip(),
        "Day": data[4],
        "Notes": data[5] if data.shape[1] > 5 else None,
    })
    df = df[df.Roll.str.lower() != "nan"].dropna(subset=["Roll"])
    df["Day"] = df.Day.apply(lambda v: "" if pd.isna(v) else str(v).strip())
    df["Notes"] = df.Notes.apply(lambda v: "" if pd.isna(v) else str(v).strip())
    return df.reset_index(drop=True)


def ta_schedule(tas: pd.DataFrame) -> dict[str, dict[str, list[str]]]:
    """weekday -> {'Morning': [names], 'Afternoon': [names]} built from every
    Duty == 'Attendance' row (Day/Notes hold the weekday and slot for those
    rows specifically - see `load_tas`)."""
    sched: dict[str, dict[str, list[str]]] = {}
    for _, r in tas[tas.Duty == "Attendance"].iterrows():
        sched.setdefault(r.Day, {}).setdefault(r.Notes, []).append(r.Name)
    return sched


def ta_for_date(sched: dict, date, slot: str) -> str | None:
    """The on-duty TA name(s) for a given date + slot ('Morning'/'Afternoon'),
    joined with ' & ' - or None if that date's weekday isn't in the roster
    (a make-up class on an off-schedule day, most likely)."""
    if date is None:
        return None
    weekday = date.strftime("%A")
    names = sched.get(weekday, {}).get(slot)
    return " & ".join(names) if names else None


def build(cohorts, out_html: str, *, course=C.COURSE,
          session=C.SESSION, tas: pd.DataFrame | None = None,
          attendance: dict | None = None, att_asof: str = "") -> str:
    """cohorts: list of dicts with keys alloc, room, cohort, label.
    tas: optional DataFrame from `load_tas` - Roll/Name/Duty/Day/Notes.
    attendance: optional {roll: {percent, present, held}} from
    `attendance_report.percent_map`; each student then also sees their own
    attendance when they look themselves up. Keyed by the same one-way hash as
    the block lookup, so the published page still carries no roll numbers in
    the clear and no names at all - a student can read their own figure only by
    already knowing their own roll number.
    """
    lookup, rooms = {}, {}
    for c in cohorts:
        room = c["room"]
        for _, s in c["alloc"].iterrows():
            lookup[roll_key(s.Roll)] = f"{room}|{s.Block}"
        counts = {b.key: int((c["alloc"].Block == b.key).sum()) for b in ROOMS[room]}
        rooms[room] = {
            "label": c.get("label", room), "cohort": c["cohort"],
            "n": len(c["alloc"]),
            "blocks": [{"k": b.key, "side": b.side, "cols": b.cols, "rows": b.rows,
                        "seated": b.core, "extra": b.extra, "n": counts[b.key],
                        "sec": b.section, "col": "#" + BLOCK_COLOUR[b.key]}
                       for b in ROOMS[room]]}
    ta_lookup = {}
    if tas is not None:
        for _, t in tas.iterrows():
            ta_lookup[roll_key(t.Roll)] = "|".join((t.Duty, t.Day, t.Notes))
    att_lookup = {}
    for roll, rec in (attendance or {}).items():
        pct, held = rec.get("percent"), rec.get("held") or 0
        if pct is None or not held:
            continue
        att_lookup[roll_key(roll)] = f"{round(pct)}|{rec.get('present', 0)}|{held}"

    payload = {"lookup": lookup, "rooms": rooms, "order": [c["room"] for c in cohorts],
               "course": course, "session": session, "ta": ta_lookup,
               "att": att_lookup, "attAsOf": att_asof}
    tpl = (Path(__file__).parent / "finder template.html").read_text(encoding="utf-8")
    Path(out_html).write_text(
        tpl.replace("__DATA__", json.dumps(payload, separators=(",", ":"), ensure_ascii=False)),
        encoding="utf-8")
    return out_html
