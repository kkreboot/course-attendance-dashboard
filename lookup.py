#!/usr/bin/env python3
"""Everything the toolkit knows about one student, in one place.

A query from a student - "where do I sit?", "which classes did I miss?",
"did you mail me?" - used to mean opening four pages and a workbook. This
module gathers the answers from the files that already hold them, and writes
nothing:

  classroom seat   out/<cohort>/allocation.csv
  exam seats       out/exams/<slug>/allocation.csv (+ exam.json)
  attendance       the Class Attendance workbook (the student's batch sheet)
  moves            out/transitions.csv
  mail             out/mail_log.csv

The dashboard's Student lookup page is a view over `student_record`.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

import attendance_report as attrep
import config as C
import exam_plan
import mailer
import seating

OUT = Path("out")


def _cohort_dirs(out: Path = OUT) -> list[Path]:
    return sorted(p for p in Path(out).glob("*")
                  if p.is_dir() and not p.name.startswith(".") and p.name != "exams"
                  and (p / "allocation.csv").exists())


def _room_of(cdir: Path) -> str:
    for q in cdir.glob("*_seating.xlsx"):
        stem = q.stem[: -len("_seating")]
        if stem in C.ROOMS:
            return stem
    return ""


def _norm(roll) -> str:
    return str(roll).strip().upper()


def directory(workbook: str | Path | None = seating.ATTENDANCE_WORKBOOK,
              out: Path = OUT) -> pd.DataFrame:
    """Every student the toolkit knows of: Roll, Name, Batch.

    The register's batch sheets first (the authority on who is in which
    batch), then anyone seated in an allocation but missing from the register,
    so a late admission who has a seat is still findable.
    """
    parts = []
    if workbook and Path(workbook).exists():
        try:
            for sh in attrep.cohort_sheets(str(workbook)):
                reg = seating.load_attendance_roster(str(workbook), sh)
                parts.append(pd.DataFrame({"Roll": reg.Roll.astype(str), "Name": reg.Name,
                                           "Batch": sh}))
        except Exception:                           # noqa: BLE001 - fall through to allocations
            pass
    for cdir in _cohort_dirs(out):
        a = pd.read_csv(cdir / "allocation.csv", dtype={"Roll": str})
        parts.append(pd.DataFrame({"Roll": a.Roll.astype(str),
                                   "Name": a.get("Name", pd.Series([""] * len(a))),
                                   "Batch": cdir.name.title()}))
    if not parts:
        return pd.DataFrame(columns=["Roll", "Name", "Batch"])
    df = pd.concat(parts, ignore_index=True).fillna("")
    df["Roll"] = df.Roll.map(_norm)
    return df.drop_duplicates("Roll", keep="first").reset_index(drop=True)


def search(query: str, people: pd.DataFrame | None = None, limit: int = 25) -> pd.DataFrame:
    """Students whose roll number or name contains `query` (case-insensitive).

    An exact roll-number match comes first; every word of a name query must
    appear, so "priya sharma" doesn't return every Priya.
    """
    people = directory() if people is None else people
    q = str(query).strip()
    if not q or people.empty:
        return people.iloc[0:0]
    ql = q.lower()
    exact = people[people.Roll.str.lower() == ql]
    by_roll = people[people.Roll.str.lower().str.contains(ql, regex=False)]
    words = ql.split()
    names = people.Name.astype(str).str.lower()
    by_name = people[names.map(lambda n: all(w in n for w in words))]
    hits = pd.concat([exact, by_roll, by_name]).drop_duplicates("Roll")
    return hits.head(limit).reset_index(drop=True)


@dataclass
class StudentRecord:
    roll: str
    name: str = ""
    batch: str = ""
    email: str = ""
    classroom: list[dict] = field(default_factory=list)   # cohort, room, block, seat
    exams: list[dict] = field(default_factory=list)       # title, when, room, block, seat
    attendance: dict | None = None                         # figures, see below
    marks: pd.DataFrame = field(default_factory=pd.DataFrame)
    moves: pd.DataFrame = field(default_factory=pd.DataFrame)
    mail: pd.DataFrame = field(default_factory=pd.DataFrame)

    @property
    def found(self) -> bool:
        return bool(self.name or self.classroom or self.exams or self.attendance)


def student_record(roll: str, *, workbook: str | Path | None = seating.ATTENDANCE_WORKBOOK,
                   out: Path = OUT, threshold: float | None = None,
                   total_classes: int | None = None,
                   excused: str = attrep.EXCUSED_MODE) -> StudentRecord:
    """Gather one student's seat, exam seats, attendance, moves and mail.

    `threshold` and `total_classes` default to the workbook's own Dashboard
    settings (then 75% / the number of dated classes), the same defaults the
    Attendance summary page uses, so the figures agree with it.
    """
    roll = _norm(roll)
    rec = StudentRecord(roll=roll)
    try:
        rec.email = mailer.student_email(roll)
    except ValueError:
        rec.email = ""

    # classroom seat
    for cdir in _cohort_dirs(out):
        a = pd.read_csv(cdir / "allocation.csv", dtype={"Roll": str})
        hit = a[a.Roll.map(_norm) == roll]
        for r in hit.itertuples():
            rec.name = rec.name or str(getattr(r, "Name", "") or "")
            rec.classroom.append({"Cohort": cdir.name.title(), "Room": _room_of(cdir),
                                  "Block": str(r.Block), "Seat": str(r.Seat)})
            rec.batch = rec.batch or str(getattr(r, "Medium", "") or cdir.name.title())

    # exam seats, newest first
    for meta in exam_plan.list_exams(Path(out) / "exams"):
        try:
            a = pd.read_csv(Path(meta["dir"]) / exam_plan.ALLOC_FILE, dtype={"Roll": str})
        except (OSError, pd.errors.ParserError):
            continue
        for r in a[a.Roll.map(_norm) == roll].itertuples():
            rec.name = rec.name or str(getattr(r, "Name", "") or "")
            rec.exams.append({"Exam": meta.get("title", meta["slug"]),
                              "When": meta.get("when", ""),
                              "Room": str(getattr(r, "Room", "")),
                              "Block": str(getattr(r, "Block", "")),
                              "Seat": str(getattr(r, "Seat", ""))})

    # attendance, from the student's own batch sheet
    if workbook and Path(workbook).exists():
        wb = str(workbook)
        try:
            settings = attrep.dashboard_blocks(wb).get("settings", {})
        except Exception:                           # noqa: BLE001 - settings are optional
            settings = {}
        thr = float(threshold if threshold is not None
                    else float(settings.get("threshold (%)") or mailer.ATTENDANCE_THRESHOLD))
        try:
            sheets = attrep.cohort_sheets(wb)
        except Exception:                           # noqa: BLE001
            sheets = []
        for sh in sheets:
            try:
                students, sessions = attrep.load(wb, sh, excused=excused)
            except (ValueError, KeyError):
                continue
            me = students[students.Roll.map(_norm) == roll]
            if me.empty:
                continue
            r = me.iloc[0]
            rec.name = rec.name or str(r.Name)
            rec.batch = sh
            held_cls = sessions[(sessions.Present + sessions.Absent + sessions.Excused) > 0]
            course_len = int(total_classes or float(settings.get("total classes (course)") or 0)
                             or len(sessions))
            remaining = max(0, course_len - len(held_cls))
            credited = int(r.Held) - int(r.Absent)
            pct = None if pd.isna(r.Percent) else float(r.Percent)
            rec.attendance = {
                "sheet": sh, "threshold": thr, "present": int(r.Present),
                "absent": int(r.Absent), "excused": int(r.Excused),
                "not_applicable": int(r.NotApplicable), "held": int(r.Held),
                "percent": pct, "classes_held": len(held_cls), "remaining": remaining,
                "asof": (held_cls.Date.max().strftime("%d %b %Y") if len(held_cls) else ""),
                "below": pct is not None and pct < thr - 1e-9,
                "to_recover": mailer.classes_to_recover(credited, int(r.Absent), thr),
                "can_miss": mailer.misses_allowed(credited, int(r.Absent), thr),
                "best_possible": mailer.best_possible_percent(credited, int(r.Absent), remaining),
                "can_reach": mailer.can_still_reach(credited, int(r.Absent), remaining, thr),
                "level": mailer.attendance_level(int(r.Absent)),
            }
            marks = attrep.load_marks(wb, sh)
            rec.marks = (marks[marks.Roll.map(_norm) == roll]
                         [["Session", "Date", "Mark"]].reset_index(drop=True))
            break

    # history
    tr = seating.load_transitions(Path(out) / "transitions.csv")
    if len(tr):
        rec.moves = tr[tr.roll.astype(str).map(_norm) == roll].reset_index(drop=True)
    log = mailer.load_log(Path(out) / "mail_log.csv")
    if len(log):
        rec.mail = (log[log.roll.astype(str).map(_norm) == roll]
                    .sort_values("timestamp", ascending=False).reset_index(drop=True))
    return rec
