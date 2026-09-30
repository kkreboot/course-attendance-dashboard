#!/usr/bin/env python3
"""The work behind the dashboard's Course setup page, kept out of the page so
it can be tested: turning the hall tables into settings and back, putting an
uploaded workbook where the dashboard looks for it, following a course-code
change with the workbook names, and reloading the toolkit so a saved change
shows up without restarting Streamlit.
"""
from __future__ import annotations

import importlib
import os
import sys
import tempfile
from pathlib import Path

import pandas as pd

import course_settings as cs

ROOM_COLS = ["Hall", "Block", "Side", "Rows", "Seats per row", "Wedge", "Section"]
GROUP_COLS = ["Hall", "Block", "Branches"]

# The workbooks setup can install, with the name each one must have. The
# dashboard finds them by these names (see config.py), so an upload is saved
# under the name, whatever it was called on the uploader's disk.
FILE_KINDS = {
    "attendance": ("Class Attendance workbook", lambda code: f"Class Attendance {code}.xlsx"),
    "roll_list": ("Department roll list", lambda code: f"student_rollList-{code}.xlsx"),
    "ta": ("TA duty workbook", lambda code: f"{code}_TA_Duty_Assignment.xlsx"),
    "exam_halls": ("Exam-hall seating plan", lambda code: "LHC Seating Plan.xlsx"),
}

# Reloaded in this order after a save: each after everything it imports.
TOOLKIT_MODULES = [
    "course_settings", "config", "history", "lockfile", "attendance_report", "mailer",
    "seating", "sheets", "reports", "handout", "finder", "ta_duty", "exam_rooms",
    "exam_plan", "posters", "attendance_pdf", "answer_showing", "checks", "lookup",
    "bundle", "course_setup",
]


# ───────────────────────────── hall tables ─────────────────────────────
def rooms_table(settings: dict) -> pd.DataFrame:
    rows = []
    for hall, blocks in settings.get("rooms", {}).items():
        for b in blocks:
            rows.append({"Hall": hall, "Block": b.get("key", ""), "Side": b.get("side", ""),
                         "Rows": int(b.get("rows", 0)), "Seats per row": int(b.get("cols", 0)),
                         "Wedge": ", ".join(str(w) for w in b.get("wedge", []) or []),
                         "Section": b.get("section", "front")})
    return pd.DataFrame(rows, columns=ROOM_COLS)


def rooms_from_table(df: pd.DataFrame) -> tuple[dict, list[str]]:
    """{hall: [block, ...]} in table order, plus anything unreadable. Rows
    with no hall and no block are ignored (the editor's empty last line)."""
    rooms: dict[str, list[dict]] = {}
    errors = []
    for i, r in enumerate(df.fillna("").to_dict("records"), start=1):
        hall, key = str(r.get("Hall", "")).strip(), str(r.get("Block", "")).strip().upper()
        if not hall and not key:
            continue
        try:
            wedge = cs.parse_int_list(r.get("Wedge", ""))
        except ValueError:
            errors.append(f"Row {i}: wedge {r.get('Wedge')!r} should be seat counts like 3, 2, 1.")
            wedge = []
        try:
            rows_n, cols_n = int(float(r.get("Rows") or 0)), int(float(r.get("Seats per row") or 0))
        except (TypeError, ValueError):
            errors.append(f"Row {i}: rows and seats per row must be numbers.")
            rows_n = cols_n = 0
        rooms.setdefault(hall, []).append({
            "key": key, "side": str(r.get("Side", "")).strip(), "rows": rows_n, "cols": cols_n,
            "wedge": wedge, "section": (str(r.get("Section", "")).strip().lower() or "front")})
    return rooms, errors


def groups_table(settings: dict) -> pd.DataFrame:
    rows = [{"Hall": hall, "Block": blk, "Branches": ", ".join(codes)}
            for hall, groups in settings.get("branch_groups", {}).items()
            for blk, codes in groups.items()]
    return pd.DataFrame(rows, columns=GROUP_COLS)


def groups_from_table(df: pd.DataFrame) -> dict:
    out: dict[str, dict[str, list[str]]] = {}
    for r in df.fillna("").to_dict("records"):
        hall, blk = str(r.get("Hall", "")).strip(), str(r.get("Block", "")).strip().upper()
        codes = cs.parse_codes(r.get("Branches", ""))
        if hall and blk and codes:
            out.setdefault(hall, {})[blk] = codes
    return out


def capacity(rooms: dict) -> pd.DataFrame:
    """Seats per hall: the full rows, the wedge seats held in reserve, total."""
    rows = []
    for hall, blocks in rooms.items():
        core = sum(int(b["rows"]) * int(b["cols"]) for b in blocks)
        extra = sum(sum(b.get("wedge", []) or []) for b in blocks)
        rows.append({"Hall": hall, "Blocks": ", ".join(b["key"] for b in blocks),
                     "Seats (full rows)": core, "Wedge seats": extra, "Total": core + extra})
    return pd.DataFrame(rows)


def halls_in_use(out: str | Path = "out") -> set[str]:
    """Halls an existing allocation was built for (`out/<cohort>/<hall>_seating.xlsx`).
    Renaming or removing one of those orphans that cohort's files."""
    used = set()
    for q in Path(out).glob("*/*_seating.xlsx"):
        used.add(q.stem[: -len("_seating")])
    return used


# ───────────────────────────── workbooks ─────────────────────────────
def check_upload(kind: str, data: bytes, filename: str = "") -> str:
    """Why an upload can't be used, or "" if it can. Reads it the way the page
    that uses it will, so a wrong file is refused here, not on exam day."""
    suffix = Path(filename).suffix.lower() or ".xlsx"
    with tempfile.TemporaryDirectory() as tmp:
        p = Path(tmp) / f"upload{suffix}"
        p.write_bytes(bytes(data))
        try:
            if kind == "attendance":
                import attendance_report as attrep
                if not attrep.register_sheets(str(p)):
                    return ("No register sheet found - a register sheet has `Roll Number` and "
                            "`Name` in its header row (row 3).")
            elif kind == "roll_list":
                import seating
                if seating.load_roster(str(p)).empty:
                    return "No students found in it."
            elif kind == "exam_halls":
                import exam_rooms
                if not exam_rooms.load_exam_rooms(str(p)):
                    return "No exam hall found in it."
            elif kind == "ta":
                import openpyxl
                openpyxl.load_workbook(p, read_only=True).close()
            else:
                return f"Unknown file kind {kind!r}."
        except Exception as e:                      # noqa: BLE001 - reported to the user
            return f"Couldn't read it: {type(e).__name__}: {e}"
    return ""


def install_file(kind: str, data: bytes, code: str, root: str | Path = ".") -> Path:
    """Write an (already checked) upload under the name the dashboard expects."""
    dest = Path(root) / FILE_KINDS[kind][1](code)
    tmp = dest.with_name(dest.name + ".tmp")
    tmp.write_bytes(bytes(data))
    os.replace(tmp, dest)
    return dest


def installed_files(code: str, root: str | Path = ".") -> dict[str, Path | None]:
    return {k: ((Path(root) / name(code)) if (Path(root) / name(code)).exists() else None)
            for k, (_, name) in FILE_KINDS.items()}


def renames_for_code(old: str, new: str, root: str | Path = ".") -> list[tuple[Path, Path]]:
    """The workbooks named after the old course code that the new code would
    no longer find, where nothing already has the new name."""
    if not old or not new or old == new:
        return []
    out = []
    for kind, (_, name) in FILE_KINDS.items():
        src, dst = Path(root) / name(old), Path(root) / name(new)
        if src != dst and src.exists() and not dst.exists():
            out.append((src, dst))
    return out


def apply_renames(pairs: list[tuple[Path, Path]]) -> list[tuple[Path, Path]]:
    done = []
    for src, dst in pairs:
        try:
            os.replace(src, dst)
            done.append((src, dst))
        except OSError:
            pass
    return done


# ───────────────────────────── reload ─────────────────────────────
def reload_toolkit() -> list[str]:
    """Re-import every toolkit module in dependency order, so values read from
    the settings at import (the course name in default arguments, the mail
    signature, the halls) pick up what was just saved. Streamlit only reruns
    the main script; without this the old course would linger until restart.
    """
    done = []
    for name in TOOLKIT_MODULES:
        mod = sys.modules.get(name)
        if mod is None:
            continue
        try:
            importlib.reload(mod)
            done.append(name)
        except Exception:                           # noqa: BLE001 - restart fixes the rest
            pass
    return done
