#!/usr/bin/env python3
"""Student mail-out for the course - seat/block transitions now, exam seating later.

Every student has an Institute mailbox derived straight from the enrolment
number that already keys every file in this project:

    B26BB1901  ->  b26bb1901@<EMAIL_DOMAIN>

so no separate address list has to be maintained (and none can go stale).
`attendance_report.py` notes that the master Class Attendance workbook grew and
lost an `Email` column more than once - addresses are derived here rather than
read from any workbook for exactly that reason.

**A transition mail names the block only, never the seat.** Classroom seating
is block-level by design (see AGENTS.md: `Seat`/`SeatRow`/`SeatCol` are a print
and attendance order, not a literal chair), so telling a student a seat number
would be telling them something untrue about where they must sit. That rule is
enforced, not just documented: `build_transition_mail` never receives a seat,
and `_no_seat_tokens` re-checks the finished body for anything shaped like a
seat id before the message is allowed out.

Exam mail is the other case: an exam seat *is* a literal chair, so
`build_exam_mail` may carry room + block + seat.

Configuration (nothing secret is ever written into this Dropbox-synced folder):

    COURSE_SMTP_HOST      e.g. smtp.gmail.com / smtp.<your-institute>
    COURSE_SMTP_PORT      587 (STARTTLS, default) or 465 (implicit SSL)
    COURSE_SMTP_USER      login, usually the full sender address
    COURSE_SMTP_PASSWORD  app password / SMTP password  ← env var only
    COURSE_SMTP_FROM      sender address        (defaults to ...USER)
    COURSE_SMTP_FROM_NAME display name         (default "<COURSE_CODE> Course Office")
    COURSE_SMTP_REPLY_TO  optional Reply-To
    COURSE_SMTP_SSL       "1" to force implicit SSL regardless of port
    COURSE_MAIL_BCC       optional address copied on every send (course office archive)

`mail_config.json` in the project folder may set any of the same keys except
the password (host/port/user/from/from_name/reply_to/bcc).

A `~/.course_dashboard_smtp.env` file (plain `KEY=value` lines) is read as a last
resort for any of the variables above - it lives in the home directory, not
in this Dropbox-synced folder, so the app password stays on the one machine
it was issued for and survives launching the dashboard from Finder or a
desktop shortcut, where a shell `export` never reaches. Precedence:
environment → `~/.course_dashboard_smtp.env` → `mail_config.json`. See README ("Gmail
SMTP").
"""
from __future__ import annotations

import csv
import json
import os
import math
import re
import smtplib
import ssl
import time
from dataclasses import dataclass, field
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid
from urllib.parse import quote, urlencode

import config as C
from pathlib import Path

INSTITUTE_DOMAIN = C.EMAIL_DOMAIN
CONFIG_FILE = "mail_config.json"
ENV_FILE = Path.home() / ".course_dashboard_smtp.env"
MAIL_LOG = Path("out") / "mail_log.csv"

# Seat ids look like "A-001" / "LHC110 B-12" - anything of that shape must not
# appear in a classroom transition mail. Roll numbers ("B26BB1901") have no
# hyphen and so don't trip this.
SEAT_TOKEN_RE = re.compile(r"\b[A-G]-\s?\d{1,3}\b")
ROLL_RE = re.compile(r"^[A-Za-z0-9]+$")
# Google displays an app password as four space-separated groups of four.
APP_PASSWORD_RE = re.compile(r"[A-Za-z0-9]{4}(?: [A-Za-z0-9]{4}){3}")


# ─────────────────────────────── addresses ───────────────────────────────
def student_email(roll: str, domain: str = INSTITUTE_DOMAIN) -> str:
    """`B26BB1901` → `b26bb1901@<EMAIL_DOMAIN>`.

    Lower-cased: Institute addresses are issued in lower case, and while SMTP
    local parts are case-sensitive in principle, a mismatched case here reads
    wrong in every To: header a student ever sees.
    """
    roll = str(roll).strip()
    if not roll or not ROLL_RE.match(roll):
        raise ValueError(f"{roll!r} is not a usable enrolment number for an address.")
    return f"{roll.lower()}@{domain}"


# ──────────────────────────────── config ─────────────────────────────────
@dataclass
class SMTPConfig:
    host: str = ""
    port: int = 587
    user: str = ""
    password: str = ""
    sender: str = ""
    sender_name: str = f"{C.COURSE_CODE} Course Office"
    reply_to: str = ""
    bcc: str = ""
    use_ssl: bool = False

    @property
    def configured(self) -> bool:
        return bool(self.host and self.sender and self.password)

    def missing(self) -> list[str]:
        gaps = []
        if not self.host:
            gaps.append("COURSE_SMTP_HOST")
        if not (self.sender or self.user):
            gaps.append("COURSE_SMTP_FROM (or COURSE_SMTP_USER)")
        if not self.password:
            gaps.append("COURSE_SMTP_PASSWORD")
        return gaps

    def from_header(self) -> str:
        return formataddr((self.sender_name, self.sender)) if self.sender_name else self.sender


def load_env_file(path: str | Path | None = None) -> dict[str, str]:
    """Parse `KEY=value` lines out of `~/.course_dashboard_smtp.env`, if it exists.

    Deliberately not `.env` in the project folder: that would sync the
    password to every machine this Dropbox folder reaches. Blank lines and
    `#` comments are skipped, surrounding quotes stripped; anything
    unparseable is ignored rather than raising, since a typo here must not
    stop the dashboard from starting.
    """
    path = Path(path if path is not None else ENV_FILE)
    if not path.exists():
        return {}
    out: dict[str, str] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[len("export "):]
        key, sep, val = line.partition("=")
        if not sep:
            continue
        val = val.strip()
        if len(val) >= 2 and val[0] == val[-1] and val[0] in "\"'":
            val = val[1:-1]
        out[key.strip()] = val
    return out


def _clean_password(pw: str) -> str:
    """Google shows an app password as `abcd efgh ijkl mnop`; people paste it
    with the spaces. Strip them only for exactly that shape - four groups of
    four alphanumerics. A real passphrase may legitimately contain spaces,
    and silently eating those would produce an auth failure no one could
    explain."""
    return pw.replace(" ", "") if APP_PASSWORD_RE.fullmatch(pw.strip()) else pw


def load_smtp_config(config_file: str | Path = CONFIG_FILE) -> SMTPConfig:
    """`mail_config.json` (non-secret keys) overlaid by environment variables.

    The password is deliberately env-only: this folder syncs through Dropbox
    across three machines, and a password sitting in a JSON file here would
    sync with it.
    """
    env_file = load_env_file()
    data: dict = {}
    path = Path(config_file)
    if path.exists():
        try:
            data = json.loads(path.read_text(encoding="utf-8")) or {}
        except json.JSONDecodeError as e:
            raise ValueError(f"{path} is not valid JSON: {e}") from e
        data.pop("password", None)

    def pick(env: str, key: str, default=""):
        return os.environ.get(env) or env_file.get(env) or data.get(key, default)

    port = int(pick("COURSE_SMTP_PORT", "port", 587) or 587)
    user = pick("COURSE_SMTP_USER", "user")
    sender = pick("COURSE_SMTP_FROM", "from") or user
    cfg = SMTPConfig(
        host=pick("COURSE_SMTP_HOST", "host"),
        port=port,
        user=user or sender,
        password=_clean_password(os.environ.get("COURSE_SMTP_PASSWORD")
                                 or env_file.get("COURSE_SMTP_PASSWORD", "")),
        sender=sender,
        sender_name=pick("COURSE_SMTP_FROM_NAME", "from_name", f"{C.COURSE_CODE} Course Office"),
        reply_to=pick("COURSE_SMTP_REPLY_TO", "reply_to"),
        bcc=pick("COURSE_MAIL_BCC", "bcc"),
        use_ssl=str(pick("COURSE_SMTP_SSL", "ssl", "")).lower() in {"1", "true", "yes"} or port == 465,
    )
    return cfg


# ──────────────────────────────── messages ───────────────────────────────
@dataclass
class Mail:
    """One composed, not-yet-sent message.

    `to` may hold several addresses for a group mail (one block's students);
    `bcc` likewise. Students always go in **Bcc** when there is more than one
    of them - a To: line of forty classmates publishes forty addresses and
    invites reply-all.
    """
    roll: str
    to: str
    subject: str
    body: str
    name: str = ""
    cc: str = ""
    bcc: str = ""
    kind: str = "generic"
    meta: dict = field(default_factory=dict)

    def preview(self) -> str:
        cc = f"Cc: {self.cc}\n" if self.cc else ""
        bcc = f"Bcc: {self.bcc}\n" if self.bcc else ""
        return f"To: {self.to}\n{cc}{bcc}Subject: {self.subject}\n\n{self.body}"


def _no_seat_tokens(text: str, what: str) -> None:
    hit = SEAT_TOKEN_RE.search(text)
    if hit:
        raise ValueError(
            f"{what} would disclose a seat id ({hit.group(0)!r}). Classroom seating is "
            "allotted by block, not by chair - send the block only."
        )


def _signature(course: str, session: str, contact: str) -> str:
    tail = f"\n{contact}" if contact else ""
    return f"\n- {course} Course Office\n{session}{tail}\n"


def build_transition_mail(roll: str, new_block: str, *, name: str = "",
                          old_block: str = "", room: str = "", old_room: str = "",
                          course: str = C.COURSE,
                          session: str = C.SESSION,
                          medium: str = "", effective: str = "",
                          contact: str = "", domain: str = INSTITUTE_DOMAIN) -> Mail:
    """Block-transition notice. Takes no seat argument, by design."""
    to = student_email(roll, domain)
    where_new = f"Block {new_block}" + (f" in {room}" if room else "")
    where_old = ""
    if old_block:
        where_old = f"Block {old_block}" + (f" in {old_room or room}" if (old_room or room) else "")

    lines = [f"Dear {name.strip() or roll},", ""]
    if where_old:
        lines.append(f"Your seating block for {course} has been changed from {where_old} "
                     f"to {where_new}.")
    else:
        lines.append(f"Your seating block for {course} is {where_new}.")
    if medium:
        lines.append(f"Batch/medium: {medium}.")
    if effective:
        lines.append(f"Effective from: {effective}.")
    lines += [
        "",
        "Seating is allotted block-wise - sit anywhere inside your block, and sign "
        "the attendance sheet for that block only. Attendance already recorded "
        "against your previous block is unaffected.",
        "",
        "Write back to this address if this does not look right.",
    ]
    body = "\n".join(lines) + _signature(course, session, contact)
    subject = f"{course} - seating block updated: Block {new_block}"

    _no_seat_tokens(body, "This transition mail")
    _no_seat_tokens(subject, "This transition subject line")
    return Mail(roll=str(roll), to=to, subject=subject, body=body, name=name,
                kind="seat_transition",
                meta={"old_block": old_block, "new_block": new_block,
                      "room": room, "old_room": old_room})


def build_exam_mail(roll: str, *, name: str = "", room: str = "", block: str = "",
                    seat: str = "", exam: str = "", when: str = "",
                    course: str = C.COURSE,
                    session: str = C.SESSION,
                    include_seat: bool = True, note: str = "",
                    contact: str = "", lookup_url: str = C.LOOKUP_URL,
                    domain: str = INSTITUTE_DOMAIN) -> Mail:
    """Exam-seating notice. An exam seat is a real chair, so it may be named -
    pass `include_seat=False` for a block-only exam venue."""
    to = student_email(roll, domain)
    import exam_rooms
    if room and exam_rooms.is_row_col_room(room):
        where = ", ".join(p for p in (
            f"Room {room}" if room else "",
            f"Seat {seat}" if (include_seat and seat) else "",
        ) if p)
    else:
        where = ", ".join(p for p in (
            f"Room {room}" if room else "",
            f"Block {block}" if block else "",
            f"Seat {seat}" if (include_seat and seat) else "",
        ) if p)

    head = f"{course} - {exam}" if exam else f"{course} exam"
    lines = [f"Dear {name.strip() or roll},", "",
             f"Your seating for {head}:", f"    {where or 'to be announced'}"]
    if when:
        lines.append(f"    {when}")
    lines += ["", "Carry your Institute ID card. Report to the venue at least fifteen "
              "minutes before the start time."]
    if lookup_url:
        lines += ["", f"You can also check your position by entering roll number on this web page: {lookup_url}"]
    if note:
        lines += ["", note]
    lines += ["", "Write back to this address if this does not look right."]
    body = "\n".join(lines) + _signature(course, session, contact)
    subject = f"{head} - seating: " + (where or "venue to be announced")
    if not include_seat:
        _no_seat_tokens(body, "This exam mail")

    return Mail(roll=str(roll), to=to, subject=subject, body=body, name=name,
                kind="exam_seating",
                meta={"room": room, "block": block,
                      "seat": seat if include_seat else "", "exam": exam})


# One list, in config.py: the names printed on the packs and the addresses
# Cc'd on the mail are the same two people, and a change should not have to be
# made twice.
DEFAULT_ATTENDANCE_CC = C.instructor_emails()

# Signed by a person, not by "the course office": these mails ask the student
# to come and talk to someone, so they name who wrote them. Kept verbatim as
# supplied, Devanagari included - the body is sent as UTF-8, and the Gmail
# compose URL percent-encodes it, so both paths carry it intact.
SIGNATURE = f"""Your Name / आपका नाम
Your Roll Number
Teaching Assistant
{C.DEPARTMENT}
{C.INSTITUTE}"""

# The attendance notices sign off shorter, in the wording the course supplied:
# these go to a hundred students at a time and read as a note from their TA,
# not as a letter from an office with an address.
ATTENDANCE_SIGNATURE = f"""Your Name
Teaching Assistant, {C.COURSE}
{C.DEPARTMENT}, {C.INSTITUTE_SHORT}"""

# The attendance requirement the notices are written against.
ATTENDANCE_THRESHOLD = 75.0


# Escalation steps, by absences. A student who has been written to at "5+" and
# then misses more classes needs telling again - but the *same* letter twice
# for the same level is noise that teaches people to ignore it. The level is
# recorded in the log's `kind`, so `unmailed_attendance` can tell "already
# warned at this level" from "has since got worse".
ATTENDANCE_LEVELS = [(5, "L1"), (9, "L2"), (13, "L3")]
LEVEL_NOTE = {
    "L2": "This is a second notice - your absences have increased since the last one.",
    "L3": "This is a final notice. Please meet the course instructor before the next class.",
}


def attendance_level(absent: int, levels=ATTENDANCE_LEVELS) -> str | None:
    """Highest level whose threshold `absent` has reached, or None."""
    hit = [name for limit, name in sorted(levels) if absent >= limit]
    return hit[-1] if hit else None


def unmailed_attendance(students: "pd.DataFrame", levels=ATTENDANCE_LEVELS,
                        path: str | Path = MAIL_LOG) -> "pd.DataFrame":
    """`students` (Roll/Absent/... from `attendance_report.load`) narrowed to
    those who have reached a level they haven't been written to at yet.

    Adds a `Level` column. Someone who slipped from L1 to L2 reappears; someone
    still sitting at the level they were warned about does not.
    """
    import pandas as pd

    df = students.copy()
    df["Level"] = [attendance_level(int(a), levels) for a in df.Absent]
    df = df[df.Level.notna()]
    if df.empty:
        return df
    log = load_log(path)
    if log.empty:
        return df.reset_index(drop=True)
    told = log[log.status.isin(["sent", "handed-off"])]
    seen = {(r.roll, r.kind) for r in told.itertuples()}
    keep = [(str(r.Roll), f"attendance_shortfall_{r.Level}") not in seen for r in df.itertuples()]
    return df[keep].reset_index(drop=True)


def _pct(value) -> str:
    """One decimal, and no trailing `.0` - "66.7%" and "75%", never "75.0%"."""
    if value is None:
        return "-"
    out = f"{float(value):.1f}"
    return out[:-2] if out.endswith(".0") else out


def _classes(n: int) -> str:
    """`1 class`, `21 classes` - a notice that reads "present for 1 classes"
    looks generated, and these go out under a person's name."""
    return f"{int(n)} class" if int(n) == 1 else f"{int(n)} classes"


def classes_to_recover(present: int, absent: int,
                       threshold: float = ATTENDANCE_THRESHOLD) -> int:
    """How many classes in a row from here put a student back on the threshold.

    Solving (p + k) / (p + a + k) >= t for k. At 75% that is k >= 3a - p: every
    class missed costs three to make up, which is the number the notice has to
    state plainly, because "attend regularly" means nothing to someone who has
    missed eleven of twenty-one.
    """
    p, a, t = int(present), int(absent), float(threshold) / 100.0
    if t >= 1:
        return -1 if a else 0
    need = (t * (p + a) - p) / (1 - t)
    return max(0, math.ceil(need - 1e-9))


def misses_allowed(present: int, absent: int,
                   threshold: float = ATTENDANCE_THRESHOLD) -> int:
    """How many classes in a row a student can miss from here and still be on
    the threshold - the other side of `classes_to_recover`.

    The largest k with p / (p + a + k) >= t, i.e. k <= p/t - p - a. Zero for a
    student already below the line (or exactly on it), so "can miss 0" means
    the next absence takes them under. The at-risk forecast on the Attendance
    summary page is this number compared with the classes coming up.
    """
    p, a, t = int(present), int(absent), float(threshold) / 100.0
    if t <= 0:
        return 10 ** 6
    return max(0, math.floor(p / t - p - a + 1e-9))


def best_possible_percent(present: int, absent: int, remaining: int) -> float | None:
    """The highest percentage still reachable if every remaining class is
    attended. None when nothing has been held and nothing remains."""
    p, a, r = int(present), int(absent), max(0, int(remaining))
    total = p + a + r
    return (p + r) / total * 100 if total else None


def can_still_reach(present: int, absent: int, remaining: int,
                    threshold: float = ATTENDANCE_THRESHOLD) -> bool:
    best = best_possible_percent(present, absent, remaining)
    return best is not None and best >= threshold - 1e-9


def build_attendance_mail(roll: str, *, name: str = "", absent: int = 0, held: int = 0,
                          present: int | None = None, percent: float | None = None,
                          remaining: int = 0, asof: str = "",
                          threshold: float = ATTENDANCE_THRESHOLD,
                          meet_by: str = "", documents_by: str = "",
                          consequence: str = "",
                          missed_limit: int = 4, level: str = "",
                          cc: str = DEFAULT_ATTENDANCE_CC,
                          course: str = C.COURSE, course_code: str = "",
                          medium: str = "",
                          note: str = "", contact: str = "",
                          signature: str = ATTENDANCE_SIGNATURE,
                          domain: str = INSTITUTE_DOMAIN) -> Mail:
    """Attendance notice, in the course's own wording. **Two templates, and
    which one a student gets is arithmetic, not judgement.**

    A student who can still reach the threshold is told exactly how: the number
    of classes in a row that gets them back (`classes_to_recover`). A student
    who cannot reach it however many they attend is not told to "attend
    regularly" as though it were still in their hands - they are told the
    highest figure still open to them and asked to meet the instructors, and
    that documents for genuine absences change the arithmetic.

    `remaining` is how many classes of the course are still to be held; with
    none given, every student reads as unrecoverable, so the caller has to say.
    `asof` labels the figures with the last class actually held.

    The wording below is the course's own - don't rewrite it to read better.
    Only the name, the numbers and the two dates vary.
    """
    to = student_email(roll, domain)
    present = held - absent if present is None else int(present)
    absent = int(absent)
    if percent is None:
        counted = present + absent
        percent = present / counted * 100 if counted else None
    code = course_code or (course.split()[0] if course else "the")
    who = f"{name.strip()} ({roll})" if name.strip() else str(roll)
    asof_bit = (f"As of {asof} ({_classes(held)} held)" if asof
                else f"After {_classes(held)} held")

    reachable = can_still_reach(present, absent, remaining, threshold)
    needed = classes_to_recover(present, absent, threshold)
    best = best_possible_percent(present, absent, remaining)

    if reachable:
        subject = f"{code} attendance below {_pct(threshold)}%: action needed"
        lines = [
            f"Dear {who},",
            "",
            f"This is a reminder about your attendance in {course}. {asof_bit}, your "
            f"attendance is {_pct(percent)}%. You were present for {_classes(present)} and "
            f"absent for {absent}. Excused absences are not counted.",
            "",
            f"The minimum requirement is {_pct(threshold)}%. You can still meet it: you "
            + ("need to attend the next class" if needed == 1 else
               f"need to attend the next {_classes(needed)} in a row")
            + f" to get back to {_pct(threshold)}%. After that, keep attending regularly "
            f"so you stay above it. {_classes(remaining)} remain"
            + ("s" if int(remaining) == 1 else "") + " in the course.",
            "",
            "If any of your absences were for medical or other valid reasons, please send "
            "supporting documents to the course instructor or me as soon as possible, so "
            "they can be recorded as excused.",
        ]
    else:
        subject = f"{code} attendance shortfall: please meet the course instructor"
        meet = f"by {meet_by}" if meet_by else "as soon as possible"
        docs = f"by {documents_by or meet_by}" if (documents_by or meet_by) else "as soon as possible"
        lines = [
            f"Dear {who},",
            "",
            f"I am writing about your attendance in {course}. {asof_bit}, your attendance "
            f"is {_pct(percent)}%. You were present for {_classes(present)} and absent for "
            f"{absent}.",
            "",
            f"Even if you attend all {remaining} remaining "
            + ("class" if int(remaining) == 1 else "classes")
            + f", the highest you can reach is {_pct(best)}%. That is below the "
            f"{_pct(threshold)}% requirement.",
            "",
            "Please:",
            "",
            "1. Attend every remaining class from now on.",
            f"2. Meet the course instructors {meet} to discuss your situation.",
            f"3. Submit supporting documents {docs} if any absences were for medical or "
            "other valid reasons. They can be recorded as excused, which changes your "
            "percentage.",
        ]
        if consequence:
            lines += ["", consequence]

    # An escalation only adds a sentence saying which notice this is.
    step = LEVEL_NOTE.get(level, "")
    if step:
        lines += ["", step]
    if note:
        lines += ["", note]
    lines += ["", "Regards,"]
    if signature:
        lines.append(signature)
    if contact:
        lines += ["", contact]
    body = "\n".join(lines) + "\n"

    kind = "attendance_shortfall" if reachable else "attendance_unreachable"
    if level:
        kind = f"attendance_shortfall_{level}"
    return Mail(roll=str(roll), to=to, subject=subject, body=body, name=name, cc=cc,
                kind=kind,
                meta={"absent": absent, "held": held, "present": present, "percent": percent,
                      "remaining": int(remaining), "threshold": float(threshold),
                      "needed": needed, "best_possible": best, "reachable": reachable,
                      "limit": missed_limit, "level": level, "medium": medium,
                      "asof": asof})


def build_group_mail(rolls, subject: str, body: str, *, label: str = "",
                     cc: str = "", bcc: str = "", to: str = "",
                     students_in: str = "to", kind: str = "group",
                     domain: str = INSTITUTE_DOMAIN) -> Mail:
    """One message to a whole block (or room).

    `students_in="to"` puts every student on the To: line, which is what the
    course asked for: the block can see it is a block notice and can reply.
    The cost is that forty classmates' addresses are visible to each other and
    a reply-all reaches all of them - `students_in="bcc"` is there for when
    that matters more.

    Instructors go in Cc and the block's TAs in Bcc, so a student replying
    reaches the course office rather than the invigilators.
    """
    if students_in not in ("to", "bcc"):
        raise ValueError("students_in must be 'to' or 'bcc'")
    addresses = [student_email(r, domain) for r in rolls]
    if not addresses:
        raise ValueError("No students in this block - nothing to send.")

    if students_in == "to":
        to_line = ",".join(addresses)
        bcc_line = bcc
    else:
        to_line = to or ""
        bcc_line = ",".join(x for x in (bcc, *addresses) if x)

    return Mail(roll=label or f"{len(addresses)} students",
                to=to_line, subject=subject, body=body, cc=cc, bcc=bcc_line,
                kind=kind, meta={"n": len(addresses), "label": label,
                                 "students_in": students_in})


def ta_email(roll: str, domain: str = INSTITUTE_DOMAIN) -> str:
    """TA rolls follow the same pattern as students' - `P24PH0999` →
    `p24ph0999@<EMAIL_DOMAIN>`."""
    return student_email(roll, domain)


def build_exam_block_mail(rolls, *, exam: str, when: str, room: str, block: str = "",
                          course: str = C.COURSE, seat_note: str = "",
                          lookup_url: str = "", cc: str = "", bcc: str = "",
                          students_in: str = "to", contact: str = "",
                          seats: dict | None = None,
                          signature: str = SIGNATURE) -> Mail:
    """The seating notice for one block of one exam.

    Names the room and block, and by default not a seat: one mail goes to the
    whole block, and a seat is per student. Where to find their own seat is
    then the last line, the lookup page if there is one and the block sheet
    otherwise.

    `seats` ({roll: seat}) prints the block's seat list in the mail instead.
    That is the honest answer for a room where the students cannot be told
    individually: eighty people in LHC 105 should not have to walk to a notice
    board to learn a number that fits in one line each. It publishes roll
    numbers to the block, which the To: line does anyway, and no names.
    """
    import exam_rooms

    where = (f"{room}, {exam_rooms.group_label(room, block)}" if str(block).strip()
             else str(room))
    lines = ["Dear students,", "",
             f"Your seating for {course} - {exam}:", f"    {where}"]
    if when:
        lines.append(f"    {when}")
    lines += ["", "Carry your Institute ID card and reach the room at least ten minutes "
              "before the start time. Sit in your own seat - seats are checked against "
              "the seating list."]
    if exam_rooms.is_row_col_room(room) and not str(block).strip():
        # this hall has no block letters: the plan at the door is a grid, rows
        # numbered down the side and columns across the top
        lines += ["", "Seats in this room are numbered by row and column, not by block: "
                  "find your row first, then count along it."]
    if seats:
        listed = [(str(r), str(seats[str(r)])) for r in rolls if str(r) in seats]
        if listed:
            lines += ["", "Your seat number:"]
            # two columns of "roll  seat", so eighty students are twenty lines
            # rather than eighty
            pairs = [f"{roll}  {seat}" for roll, seat in listed]
            width = max(len(x) for x in pairs) + 4
            for i in range(0, len(pairs), 2):
                row = pairs[i:i + 2]
                lines.append("    " + row[0].ljust(width) + (row[1] if len(row) > 1 else ""))
    if lookup_url:
        lines += ["", f"You can also check your position by entering roll number on this web page: {lookup_url}"]
    elif seat_note and not seats:
        lines += ["", seat_note]
    if C.COURSE_INSTRUCTORS:
        lines += ["", f"Course instructors: {C.instructor_names()}"]
    lines += ["", "Best Regards,"]
    if signature:
        lines.append(signature)
    if contact:
        lines += ["", contact]

    subject = f"{course} - {exam}: seating for {where}"
    return build_group_mail(rolls, subject, "\n".join(lines) + "\n",
                            label=where, cc=cc, bcc=bcc, students_in=students_in,
                            kind="exam_block_seating")


# ─────────────────────── hand off to a mail client ───────────────────────
# Institute Google Workspace accounts commonly have app passwords disabled by
# the admin, which leaves SMTP with no way to authenticate. The way round it
# is not to send at all: pre-fill a Gmail compose window in the browser,
# already signed in as that account, and let the human press Send. Nothing is
# stored, nothing authenticates, and the mail genuinely comes from the
# institute address.

# Gmail's compose URL has no documented length cap, but browsers and Google's
# own front end start truncating past roughly 2 kB, and a silently truncated
# Bcc list would drop students without saying so. Batches are chunked to stay
# well under that.
GMAIL_URL_BUDGET = 1800


def gmail_compose_url(to: str = "", subject: str = "", body: str = "",
                      cc: str = "", bcc: str = "", account: str = "") -> str:
    """A URL that opens Gmail's compose window with everything filled in.

    `account` picks which signed-in Google account composes it - either an
    index (`0`, `1`, … the order they were added to the browser) or the full
    address. Getting this wrong on a browser with several accounts signed in
    is the usual reason a compose window opens as the wrong sender.
    """
    base = "https://mail.google.com/mail/"
    acct = account.strip()
    if acct.isdigit():
        base += f"u/{acct}/"
    params = {"view": "cm", "fs": "1"}
    if acct and not acct.isdigit():
        params["authuser"] = acct
    for k, v in (("to", to), ("cc", cc), ("bcc", bcc), ("su", subject), ("body", body)):
        if v:
            params[k] = v
    return base + "?" + urlencode(params, quote_via=quote)


def mailto_url(to: str = "", subject: str = "", body: str = "", bcc: str = "") -> str:
    """Same handoff for whatever the machine's default mail client is -
    Outlook, Apple Mail, Thunderbird - for anyone not composing in Gmail."""
    params = {k: v for k, v in (("subject", subject), ("body", body), ("bcc", bcc)) if v}
    return f"mailto:{quote(to)}?" + urlencode(params, quote_via=quote)


def _chunk_link(chunk: list, subject: str, body: str, to: str,
                account: str, cc: str = "") -> tuple[str, str, list]:
    bcc = ",".join(m.to for m in chunk)
    label = f"{len(chunk)} recipients" if len(chunk) > 1 else chunk[0].to
    return (label, gmail_compose_url(to=to, subject=subject, body=body, bcc=bcc,
                                     cc=cc, account=account), chunk)


def compose_links(mails: list, account: str = "", to: str = "",
                  budget: int = GMAIL_URL_BUDGET, group: bool = True) -> list:
    """Group identical messages into as few compose windows as possible.

    Messages whose subject and body match exactly (a whole block being told
    the same thing, addressed generically) become one compose with every
    recipient in **Bcc** - Bcc, not To, so no student sees another's address.
    Anything personalised stays one window per student. Returns
    `(label, url, mails_covered)` per window, chunked so no URL gets long
    enough for a browser to truncate the recipient list.

    `to` is the visible To: address - worth putting the course office there,
    since a Bcc-only mail with an empty To trips some spam filters.
    """
    if not group:
        # One window each, addressed To the student - for anything carrying that
        # student's own data (attendance figures), where a shared Bcc window
        # would be the wrong shape even though the addresses stay hidden.
        return [(m.to or m.roll,
                 gmail_compose_url(to=m.to, subject=m.subject, body=m.body,
                                   cc=m.cc, bcc=m.bcc, account=account), [m])
                for m in mails]

    groups: dict[tuple[str, str, str], list] = {}
    for m in mails:
        groups.setdefault((m.subject, m.body, m.cc), []).append(m)

    out: list = []
    for (subject, body, cc), members in groups.items():
        if len(members) == 1 and not to:
            m = members[0]
            out.append((m.to, gmail_compose_url(to=m.to, subject=subject, body=body,
                                                cc=cc, account=account), [m]))
            continue
        chunk = []
        for m in members:
            # Measure the real URL rather than estimating from address
            # lengths - percent-encoding makes an estimate drift over the
            # budget, and the whole point of the budget is that it holds.
            if chunk and len(_chunk_link(chunk + [m], subject, body, to,
                                         account, cc)[1]) > budget:
                out.append(_chunk_link(chunk, subject, body, to, account, cc))
                chunk = []
            chunk.append(m)
        if chunk:
            out.append(_chunk_link(chunk, subject, body, to, account, cc))
    return out


def log_handoff(mails: list, via: str = "gmail-compose",
                path: str | Path = MAIL_LOG) -> Path:
    """Record messages handed to a mail client. Marked `handed-off`, never
    `sent`: this toolkit only put the compose window on screen - it cannot
    know whether the human pressed Send."""
    stamp = datetime.now().isoformat(timespec="seconds")
    return log_sends([dict(timestamp=stamp, kind=m.kind, roll=m.roll, to=m.to,
                           subject=m.subject, status="handed-off", detail=via)
                      for m in mails], path)


# ───────────────────────────────── sending ───────────────────────────────
def _to_email_message(mail: Mail, cfg: SMTPConfig) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = cfg.from_header()
    msg["To"] = mail.to
    if mail.cc:
        msg["Cc"] = mail.cc
    if mail.bcc:
        # smtplib.send_message strips Bcc from the sent headers but still
        # delivers to it, which is exactly what a class mail-out needs.
        msg["Bcc"] = mail.bcc
    msg["Subject"] = mail.subject
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=(cfg.sender.split("@")[-1] or None))
    if cfg.reply_to:
        msg["Reply-To"] = cfg.reply_to
    if cfg.bcc:
        msg["Bcc"] = ",".join(x for x in (mail.bcc, cfg.bcc) if x)
    msg.set_content(mail.body)
    return msg


def _connect(cfg: SMTPConfig, timeout: int = 30):
    ctx = ssl.create_default_context()
    if cfg.use_ssl:
        srv = smtplib.SMTP_SSL(cfg.host, cfg.port, timeout=timeout, context=ctx)
    else:
        srv = smtplib.SMTP(cfg.host, cfg.port, timeout=timeout)
        srv.ehlo()
        srv.starttls(context=ctx)
        srv.ehlo()
    srv.login(cfg.user or cfg.sender, cfg.password)
    return srv


def load_log(path: str | Path = MAIL_LOG) -> "pd.DataFrame":
    """The send log as a frame - empty but correctly shaped if nothing has
    been mailed yet."""
    import pandas as pd

    cols = ["timestamp", "kind", "roll", "to", "subject", "status", "detail"]
    path = Path(path)
    if not path.exists():
        return pd.DataFrame(columns=cols)
    df = pd.read_csv(path, dtype=str).fillna("")
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    return df[cols]


def unmailed(transitions: "pd.DataFrame", kind_prefix: str = "seat_transition",
             path: str | Path = MAIL_LOG) -> "pd.DataFrame":
    """The transitions nobody has been told about yet.

    A student can move more than once in a term, so this compares *times*, not
    just roll numbers: an earlier mail about an earlier move does not cover a
    later one. `handed-off` counts as told - the Gmail window was opened, and
    re-listing it would mean mailing the same person twice for one move -
    while a `refused`/`failed` row does not.
    """
    if transitions.empty:
        return transitions
    log = load_log(path)
    if log.empty:
        return transitions.reset_index(drop=True)
    told = log[log.kind.str.startswith(kind_prefix)
               & log.status.isin(["sent", "handed-off"])]
    if told.empty:
        return transitions.reset_index(drop=True)
    latest = told.groupby("roll").timestamp.max()
    keep = [t.timestamp > latest.get(str(t.roll), "") for t in transitions.itertuples()]
    return transitions[keep].reset_index(drop=True)


def log_sends(results: list[dict], path: str | Path = MAIL_LOG) -> Path:
    """Append one row per attempt to `out/mail_log.csv` - the only record that
    a student was told anything, and the thing to check before re-sending."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    cols = ["timestamp", "kind", "roll", "to", "subject", "status", "detail"]
    new = not path.exists()
    with path.open("a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        if new:
            w.writeheader()
        for r in results:
            w.writerow({c: r.get(c, "") for c in cols})
    return path


def send_mails(mails: list[Mail], cfg: SMTPConfig | None = None, *, dry_run: bool = True,
               pause: float = 0.4, log: bool = True) -> list[dict]:
    """Send (or, by default, only rehearse) a batch, one row of outcome each.

    Never raises for a single bad recipient - one refused address must not
    abort the rest of a batch. A connection/auth failure does raise, since
    nothing after it could succeed either.
    """
    cfg = cfg or load_smtp_config()
    stamp = datetime.now().isoformat(timespec="seconds")
    results: list[dict] = []

    if dry_run:
        for m in mails:
            results.append(dict(timestamp=stamp, kind=m.kind, roll=m.roll, to=m.to,
                                subject=m.subject, status="dry-run", detail=""))
        return results

    gaps = cfg.missing()
    if gaps:
        raise RuntimeError("SMTP is not configured - set " + ", ".join(gaps) +
                           " (see the docstring in mailer.py).")

    srv = _connect(cfg)
    try:
        for i, m in enumerate(mails):
            try:
                srv.send_message(_to_email_message(m, cfg))
                results.append(dict(timestamp=datetime.now().isoformat(timespec="seconds"),
                                    kind=m.kind, roll=m.roll, to=m.to, subject=m.subject,
                                    status="sent", detail=""))
            except smtplib.SMTPRecipientsRefused as e:
                results.append(dict(timestamp=datetime.now().isoformat(timespec="seconds"),
                                    kind=m.kind, roll=m.roll, to=m.to, subject=m.subject,
                                    status="refused", detail=str(e)))
            except smtplib.SMTPException as e:
                results.append(dict(timestamp=datetime.now().isoformat(timespec="seconds"),
                                    kind=m.kind, roll=m.roll, to=m.to, subject=m.subject,
                                    status="failed", detail=str(e)))
            if pause and i < len(mails) - 1:
                time.sleep(pause)
    finally:
        with_quit = getattr(srv, "quit", None)
        if with_quit:
            try:
                with_quit()
            except smtplib.SMTPException:
                pass

    if log:
        log_sends(results)
    return results


def smtp_check(cfg: SMTPConfig | None = None) -> tuple[bool, str]:
    """Connect + authenticate and hang up - for the dashboard's 'Test connection'."""
    cfg = cfg or load_smtp_config()
    gaps = cfg.missing()
    if gaps:
        return False, "Not configured: " + ", ".join(gaps)
    try:
        srv = _connect(cfg)
        srv.quit()
        return True, f"Authenticated as {cfg.user or cfg.sender} on {cfg.host}:{cfg.port}."
    except Exception as e:                     # noqa: BLE001 - surfaced verbatim in the UI
        return False, f"{type(e).__name__}: {e}"



def write_env_file(values: dict[str, str], path: str | Path | None = None) -> Path:
    """Write `~/.course_dashboard_smtp.env` at mode 600, preserving any keys already
    there that this run didn't ask about."""
    path = Path(path if path is not None else ENV_FILE)
    existing = load_env_file(path)
    existing.update({k: v for k, v in values.items() if v})
    body = "\n".join(f"{k}={v}" for k, v in existing.items())
    path.write_text("# Course dashboard mail settings - read by mailer.py. Keep this file\n"
                    "# out of the project folder: that folder syncs via Dropbox.\n"
                    + body + "\n", encoding="utf-8")
    path.chmod(0o600)
    return path


def _configure() -> int:
    """Interactive setup: prompts for the SMTP settings, reads the password
    with echo off, writes `~/.course_dashboard_smtp.env`. The password goes from the
    keyboard into the file - never onto the command line, so it can't land in
    shell history or a process list."""
    from getpass import getpass

    current = load_env_file()

    def ask(key: str, prompt: str, default: str = "") -> str:
        default = current.get(key, default)
        got = input(f"{prompt} [{default}]: ").strip()
        return got or default

    print(f"Writing {ENV_FILE} - press Enter to keep the value in brackets.\n")
    host = ask("COURSE_SMTP_HOST", "SMTP host", "smtp.gmail.com")
    port = ask("COURSE_SMTP_PORT", "Port (587 STARTTLS, 465 SSL)", "587")
    user = ask("COURSE_SMTP_USER", "Account to log in as (the full address)")
    sender = ask("COURSE_SMTP_FROM", "From address", user)
    name = ask("COURSE_SMTP_FROM_NAME", "From display name", f"{C.COURSE_CODE} Course Office")
    bcc = ask("COURSE_MAIL_BCC", "Bcc every mail to (blank for none)")
    pw = getpass("App password (input hidden; Enter keeps the stored one): ").strip()

    if not pw and not current.get("COURSE_SMTP_PASSWORD"):
        print("No password given and none stored - nothing would send. Aborted.")
        return 1

    path = write_env_file({
        "COURSE_SMTP_HOST": host, "COURSE_SMTP_PORT": port,
        "COURSE_SMTP_USER": user, "COURSE_SMTP_FROM": sender,
        "COURSE_SMTP_FROM_NAME": name, "COURSE_MAIL_BCC": bcc,
        **({"COURSE_SMTP_PASSWORD": _clean_password(pw)} if pw else {}),
    })
    print(f"\nWrote {path} (mode 600).")
    ok, detail = smtp_check()
    print(("Connection OK - " if ok else "Connection FAILED - ") + detail)
    if ok:
        print("\nNow send yourself one real message:\n"
              "    python3 mailer.py --to you@example.com --send")
    return 0 if ok else 1


# ───────────────────────────── self-test CLI ─────────────────────────────
def _main(argv: list[str] | None = None) -> int:
    """`python mailer.py --to you@example.edu --send` - one sample message to
    one address, so SMTP can be proven working before any student is mailed.
    Composes only unless `--send` is given."""
    import argparse

    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--configure", action="store_true",
                    help=f"interactively write {ENV_FILE} (password prompt is hidden)")
    ap.add_argument("--to",
                    help="where the test goes - your own address, not a student's")
    ap.add_argument("--kind", choices=["transition", "exam"], default="transition")
    ap.add_argument("--send", action="store_true",
                    help="actually send it (default: compose and print only)")
    ap.add_argument("--roll", default="B26XX0000",
                    help="roll number to name in the sample text (default: an "
                         "obviously fake one)")
    a = ap.parse_args(argv)
    if a.configure:
        return _configure()
    if not a.to:
        ap.error("--to is required (or use --configure)")

    if a.kind == "transition":
        mail = build_transition_mail(a.roll, "C", name="Test Recipient", old_block="A",
                                     room="LHC110", effective="the next class")
    else:
        mail = build_exam_mail(a.roll, name="Test Recipient", room="LHC110", block="B",
                               seat="B-012", exam="Quiz 1 (test mail)",
                               when="date & time here")
    mail.to = a.to            # the point of a test is that it reaches *you*
    mail.kind += "_test"

    cfg = load_smtp_config()
    print(mail.preview())
    print("-" * 60)
    gaps = cfg.missing()
    if gaps:
        print("SMTP not configured - missing: " + ", ".join(gaps))
        print(f"Put them in {ENV_FILE} (see README, 'Gmail SMTP').")
        return 1
    print(f"Sending as {cfg.from_header()} via {cfg.host}:{cfg.port}"
          + (" [SSL]" if cfg.use_ssl else " [STARTTLS]"))
    ok, detail = smtp_check(cfg)
    print(("connection OK - " if ok else "connection FAILED - ") + detail)
    if not ok:
        return 1
    if not a.send:
        print("Composed only. Re-run with --send to actually send it.")
        return 0
    for r in send_mails([mail], cfg, dry_run=False):
        print(f"{r['status']}: {r['to']} {r['detail']}")
    print(f"Logged to {MAIL_LOG}.")
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
