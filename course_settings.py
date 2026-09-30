#!/usr/bin/env python3
"""The course's own details, entered once and read everywhere.

`course_settings.json` in the project folder holds everything that is true for
the whole course: code, title, session and dates, the institute and mail
domain, the instructors, the TA who signs the mail, the attendance rules and
the classroom halls. The dashboard's **Course setup** page writes it;
`config.py` and `mailer.py` read it at import, so every page, sheet, poster,
PDF and mail picks the values up without anyone editing Python.

With no file the demo values in `DEFAULTS` apply, which is what the public
demo and the test suite run on. A file only needs the keys that differ: it
is merged over the defaults section by section.

This module imports nothing from the toolkit (config imports *it*), and it
never raises on a bad file - a half-synced Dropbox copy must not take the
dashboard down. `problems()` says what is wrong instead.
"""
from __future__ import annotations

import copy
import json
import os
import re
from datetime import date
from pathlib import Path

# COURSE_SETTINGS_FILE points elsewhere - the test suite uses it to run on the
# demo course whatever the folder's own settings say.
SETTINGS_FILE = Path(os.environ.get("COURSE_SETTINGS_FILE") or "course_settings.json")
VERSION = 1

# The demo course. Also the shape of the file: a key not listed here is ignored.
DEFAULTS: dict = {
    "version": VERSION,
    "course": {
        "code": "DEMO101",
        "title": "Demo Course",
        "session": "AY 2026-27 Sem 1",
        "start_date": "",          # ISO dates, optional
        "end_date": "",
        "total_classes": 39,       # planned classes in the whole course
    },
    "institute": {
        "name": "Example Institute of Technology",
        "short": "EIT",
        "department": "Department of Physics",
        "email_domain": "example.edu",
        "lookup_url": "",
    },
    "instructors": [
        {"name": "Dr. Instructor One", "email": "instructor.one@example.edu"},
        {"name": "Dr. Instructor Two", "email": "instructor.two@example.edu"},
    ],
    "ta": {                        # who signs the mail
        "name": "Your Name",
        "name_local": "आपका नाम",   # second-language name on bilingual mail; blank to omit
        "roll": "Your Roll Number",
        "email": "",
        "role": "Teaching Assistant",
        "contact": "",             # optional line under every mail (office hours, phone)
    },
    "attendance": {
        "threshold": 75.0,         # % below which a student is short
        "benchmark": 85.0,         # a class under this turnout is marked in the report
        "levels": [5, 9, 13],      # absences that trigger notice 1, 2, 3
        "excused": "exclude",      # an E: "exclude" from the denominator, or count "present"
    },
    "mail": {
        "sender": "",              # From: address for SMTP; the password is never stored here
    },
    # Classroom halls, block by block. `wedge` is the tapering rows behind the
    # last full row, seats per row; `section` is front or rear.
    "rooms": {
        "LHC110": [
            {"key": "A", "side": "Front left", "rows": 6, "cols": 6, "wedge": [], "section": "front"},
            {"key": "B", "side": "Front centre-L", "rows": 7, "cols": 6, "wedge": [3, 2, 2, 1], "section": "front"},
            {"key": "C", "side": "Front centre-R", "rows": 7, "cols": 6, "wedge": [4, 3, 2, 2, 1], "section": "front"},
            {"key": "D", "side": "Front right", "rows": 6, "cols": 6, "wedge": [], "section": "front"},
            {"key": "E", "side": "Rear left", "rows": 14, "cols": 10, "wedge": [], "section": "rear"},
            {"key": "F", "side": "Rear centre", "rows": 13, "cols": 10, "wedge": [], "section": "rear"},
            {"key": "G", "side": "Rear right", "rows": 14, "cols": 10, "wedge": [], "section": "rear"},
        ],
        "LHC2 101": [
            {"key": "A", "side": "Left", "rows": 11, "cols": 4, "wedge": [], "section": "front"},
            {"key": "B", "side": "Centre-Left", "rows": 11, "cols": 4, "wedge": [], "section": "front"},
            {"key": "C", "side": "Centre-Right", "rows": 11, "cols": 4, "wedge": [], "section": "front"},
            {"key": "D", "side": "Right", "rows": 11, "cols": 4, "wedge": [], "section": "front"},
        ],
    },
    # Which branch codes sit together in which block, per room (optional).
    "branch_groups": {
        "LHC110": {"A": ["BB", "CI", "MA"], "B": ["PH", "MT", "ME"], "C": ["CM", "EE", "CY"],
                   "D": ["AE", "CS"], "F": ["CH", "EC"]},
        "LHC2 101": {"A": ["CS", "ME"], "B": ["EC", "AE", "CH"], "C": ["CI", "EE", "BB"],
                     "D": ["CM", "MA", "PH", "CY"]},
    },
    "saved": "",
}

# Values that mean "not set up yet" - the health check flags them.
DEMO_MARKERS = {
    ("institute", "email_domain"): "example.edu",
    ("ta", "name"): "Your Name",
    ("course", "code"): "DEMO101",
}

BLOCK_KEY_RE = re.compile(r"^[A-Z]$")
CODE_RE = re.compile(r"^[A-Za-z0-9]+$")
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
DOMAIN_RE = re.compile(r"^[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$")


def _merge(base: dict, over: dict) -> dict:
    """`over` on top of `base`, one level of sections deep. Lists and the
    rooms/branch tables are replaced whole - a hall is not merged block by
    block with the demo's."""
    out = copy.deepcopy(base)
    for k, v in (over or {}).items():
        if k not in out:
            continue
        if isinstance(out[k], dict) and isinstance(v, dict) and k not in ("rooms", "branch_groups"):
            out[k] = {**out[k], **{kk: vv for kk, vv in v.items() if kk in out[k]}}
        else:
            out[k] = copy.deepcopy(v)
    return out


def exists(path: str | Path = SETTINGS_FILE) -> bool:
    return Path(path).exists()


def read_raw(path: str | Path = SETTINGS_FILE) -> dict:
    """The file as written, or {} if it is missing or unreadable."""
    p = Path(path)
    if not p.exists():
        return {}
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (json.JSONDecodeError, OSError, UnicodeDecodeError):
        return {}


def load(path: str | Path = SETTINGS_FILE) -> dict:
    """The effective settings: the file merged over `DEFAULTS`."""
    return _merge(DEFAULTS, read_raw(path))


def problems(s: dict) -> list[str]:
    """Everything that would make a page or a mail wrong, as sentences.
    Empty means the settings are usable."""
    out = []
    c, i = s.get("course", {}), s.get("institute", {})
    if not CODE_RE.match(str(c.get("code", ""))):
        out.append("Course code must be letters and digits only (it is used in file names).")
    if not str(c.get("title", "")).strip():
        out.append("Course title is empty.")
    if not DOMAIN_RE.match(str(i.get("email_domain", ""))):
        out.append("Student email domain should look like `iitj.ac.in`.")
    for field in ("start_date", "end_date"):
        v = str(c.get(field, "") or "")
        if v:
            try:
                date.fromisoformat(v)
            except ValueError:
                out.append(f"Course {field.replace('_', ' ')} {v!r} is not a date (YYYY-MM-DD).")
    if c.get("start_date") and c.get("end_date") and str(c["end_date"]) < str(c["start_date"]):
        out.append("Course end date is before its start date.")
    try:
        if int(c.get("total_classes", 0)) < 1:
            out.append("Total classes planned must be at least 1.")
    except (TypeError, ValueError):
        out.append("Total classes planned must be a number.")
    if not s.get("instructors"):
        out.append("Add at least one course instructor.")
    for ins in s.get("instructors", []):
        if not str(ins.get("name", "")).strip():
            out.append("An instructor has no name.")
        if ins.get("email") and not EMAIL_RE.match(str(ins["email"])):
            out.append(f"Instructor email {ins['email']!r} doesn't look like an address.")
    ta_email = s.get("ta", {}).get("email")
    if ta_email and not EMAIL_RE.match(str(ta_email)):
        out.append(f"TA email {ta_email!r} doesn't look like an address.")
    sender = s.get("mail", {}).get("sender")
    if sender and not EMAIL_RE.match(str(sender)):
        out.append(f"Mail sender {sender!r} doesn't look like an address.")
    a = s.get("attendance", {})
    try:
        if not 0 < float(a.get("threshold", 0)) <= 100:
            out.append("Attendance threshold must be between 0 and 100.")
        if not 0 <= float(a.get("benchmark", 0)) <= 100:
            out.append("Class benchmark must be between 0 and 100.")
    except (TypeError, ValueError):
        out.append("Attendance threshold and benchmark must be numbers.")
    lv = a.get("levels", [])
    if not (isinstance(lv, list) and lv and all(isinstance(x, int) and x > 0 for x in lv)
            and lv == sorted(set(lv))):
        out.append("Notice levels must be increasing whole numbers of absences, e.g. 5, 9, 13.")
    if a.get("excused") not in ("exclude", "present"):
        out.append("Excused absences must be 'exclude' or 'present'.")
    rooms = s.get("rooms", {})
    if not rooms:
        out.append("Add at least one classroom hall.")
    for room, blocks in rooms.items():
        if not str(room).strip():
            out.append("A hall has no name.")
        keys = [str(b.get("key", "")) for b in blocks]
        if not blocks:
            out.append(f"Hall {room!r} has no blocks.")
        if len(keys) != len(set(keys)):
            out.append(f"Hall {room!r} has two blocks with the same letter.")
        for b in blocks:
            if not BLOCK_KEY_RE.match(str(b.get("key", ""))):
                out.append(f"Hall {room!r}: block key {b.get('key')!r} must be one capital letter.")
            try:
                if int(b.get("rows", 0)) < 1 or int(b.get("cols", 0)) < 1:
                    out.append(f"Hall {room!r} block {b.get('key')}: rows and seats per row must be at least 1.")
                if any(int(w) < 1 for w in b.get("wedge", [])):
                    out.append(f"Hall {room!r} block {b.get('key')}: wedge rows need at least 1 seat each.")
            except (TypeError, ValueError):
                out.append(f"Hall {room!r} block {b.get('key')}: rows, seats and wedge must be numbers.")
            if b.get("section") not in ("front", "rear"):
                out.append(f"Hall {room!r} block {b.get('key')}: section must be front or rear.")
    for room, groups in s.get("branch_groups", {}).items():
        if room not in rooms:
            out.append(f"Branch grouping names hall {room!r}, which isn't set up.")
            continue
        keys = {b.get("key") for b in rooms[room]}
        for blk in groups:
            if blk not in keys:
                out.append(f"Branch grouping: hall {room!r} has no block {blk!r}.")
    return out


def demo_leftovers(s: dict) -> list[str]:
    """Fields still holding the demo's value - fine for the demo, wrong for a course."""
    return [f"{sec}.{key}" for (sec, key), demo in DEMO_MARKERS.items()
            if str(s.get(sec, {}).get(key, "")) == demo]


def save(s: dict, path: str | Path = SETTINGS_FILE) -> Path:
    """Write `s` (validated by the caller), atomically: a temporary file then a
    rename, so a Dropbox sync never sees half a file."""
    from datetime import datetime

    p = Path(path)
    data = _merge(DEFAULTS, s)
    data["version"] = VERSION
    data["saved"] = datetime.now().isoformat(timespec="seconds")
    tmp = p.with_name(p.name + ".tmp")
    tmp.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, p)
    return p


# ─────────────────────────── helpers for the form ───────────────────────────
def parse_int_list(text: str) -> list[int]:
    """'3, 2, 2, 1' -> [3, 2, 2, 1]; blanks ignored. Raises ValueError on junk."""
    return [int(x) for x in re.split(r"[,\s]+", str(text or "").strip()) if x]


def parse_codes(text: str) -> list[str]:
    """'bb, CI  ma' -> ['BB', 'CI', 'MA']."""
    return [x.upper() for x in re.split(r"[,\s]+", str(text or "").strip()) if x]


def parse_people(text: str) -> list[dict]:
    """One per line, 'Name, email' (or 'Name <email>'); email optional."""
    out = []
    for line in str(text or "").splitlines():
        line = line.strip()
        if not line:
            continue
        m = re.match(r"^(.*?)\s*<([^>]+)>\s*$", line)
        if m:
            name, email = m.group(1), m.group(2)
        elif "," in line:
            name, email = line.rsplit(",", 1)
        else:
            name, email = line, ""
        out.append({"name": name.strip(), "email": email.strip()})
    return out


def format_people(people: list[dict]) -> str:
    return "\n".join(f"{p.get('name', '')}, {p.get('email', '')}".rstrip(", ") for p in people)
