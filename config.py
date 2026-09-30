"""Everything course- and room-specific lives here.

The values come from `course_settings.json`, which the dashboard's **Course
setup** page writes (see `course_settings.py`). With no such file, the demo
course below applies. Edit the file through the page, not this module: no
other file names the course.
"""
from dataclasses import dataclass

import course_settings as _cs

SETTINGS = _cs.load()
_course, _inst, _ta, _att = (SETTINGS["course"], SETTINGS["institute"],
                             SETTINGS["ta"], SETTINGS["attendance"])

# ── The course ──────────────────────────────────────────────────────────────
COURSE_CODE = str(_course["code"]).strip() or "COURSE"
COURSE_TITLE = str(_course["title"]).strip()
COURSE = f"{COURSE_CODE} {COURSE_TITLE}".strip()  # printed on every sheet, page and mail
SESSION = str(_course["session"]).strip()
COURSE_START = str(_course.get("start_date") or "")
COURSE_END = str(_course.get("end_date") or "")
TOTAL_CLASSES = int(_course.get("total_classes") or 0)

INSTITUTE = str(_inst["name"]).strip()
INSTITUTE_SHORT = str(_inst["short"]).strip()
DEPARTMENT = str(_inst["department"]).strip()
# Student and TA addresses are derived, never stored: <roll>@EMAIL_DOMAIN.
EMAIL_DOMAIN = str(_inst["email_domain"]).strip()
# Where the published "find your block / seat" page lives, if anywhere.
LOOKUP_URL = str(_inst.get("lookup_url") or "").strip()

# Who signs the mail (mailer.SIGNATURE / ATTENDANCE_SIGNATURE are built from these).
TA_NAME = str(_ta.get("name") or "").strip()
TA_NAME_LOCAL = str(_ta.get("name_local") or "").strip()
TA_ROLL = str(_ta.get("roll") or "").strip()
TA_EMAIL = str(_ta.get("email") or "").strip()
TA_ROLE = str(_ta.get("role") or "Teaching Assistant").strip()
TA_CONTACT = str(_ta.get("contact") or "").strip()
MAIL_SENDER = str(SETTINGS["mail"].get("sender") or "").strip()

# Attendance rules.
ATTENDANCE_THRESHOLD = float(_att["threshold"])
ATTENDANCE_BENCHMARK = float(_att["benchmark"])
ATTENDANCE_LEVELS = [int(x) for x in _att["levels"]]
EXCUSED_MODE = _att["excused"] if _att.get("excused") in ("exclude", "present") else "exclude"

# The workbooks the dashboard reads from the project folder. The names must
# keep "Class Attendance" and "TA_Duty" in them: the pages find them by that.
ATTENDANCE_WORKBOOK = f"Class Attendance {COURSE_CODE}.xlsx"
TA_WORKBOOK = f"{COURSE_CODE}_TA_Duty_Assignment.xlsx"
ROLL_LIST = f"student_rollList-{COURSE_CODE}.xlsx"
EXAM_HALLS_WORKBOOK = "LHC Seating Plan.xlsx"

APP_NAME = f"{COURSE_CODE} Dashboard"

@dataclass
class Block:
    key: str
    side: str
    rows: int          # full-width rows
    cols: int          # seats per row
    wedge: tuple = ()  # taper behind the last full row, seats per wedge row
    section: str = "front"

    @property
    def core(self) -> int:
        return self.rows * self.cols

    @property
    def extra(self) -> int:
        return sum(self.wedge)

    @property
    def total(self) -> int:
        return self.core + self.extra


def _blocks(spec: list[dict]) -> list[Block]:
    return [Block(str(b["key"]), str(b.get("side", "")), int(b["rows"]), int(b["cols"]),
                  tuple(int(w) for w in b.get("wedge", []) or ()),
                  b.get("section", "front")) for b in spec]


# Classroom halls, block by block (Course setup page -> Halls).
ROOMS = {str(name): _blocks(spec) for name, spec in SETTINGS["rooms"].items()}
# The demo's two halls under their old names, for anything that imported them.
LHC110 = ROOMS.get("LHC110", [])
LHC2_101 = ROOMS.get("LHC2 101", [])

# Who the course belongs to. Printed on anything a student or an invigilator
# reads, and the source of the Cc list on every mail the toolkit sends, so the
# names and the addresses cannot drift apart.
COURSE_INSTRUCTORS = [(str(p.get("name", "")).strip(), str(p.get("email", "")).strip())
                      for p in SETTINGS["instructors"] if str(p.get("name", "")).strip()]


def instructor_names(sep: str = ", ") -> str:
    return sep.join(name for name, _ in COURSE_INSTRUCTORS)


def instructor_emails(sep: str = ",") -> str:
    return sep.join(email for _, email in COURSE_INSTRUCTORS if email)

# Optional branch-grouped seating: which
# 2-letter department codes (see seating.py's `branch_code`) sit together in
# which block, in room order. A block with no entry here just isn't part of
# the grouping - it's filled by the catch-all pass in `allocate_grouped`
# (backlog students, and any branch not listed for any block). Only current
# B26 branches are grouped; backlog students are never split by branch, per
# the arrangement sheet ("Backlogs - all branches" as one lump total).
BRANCH_GROUPS = {room: {blk: list(codes) for blk, codes in groups.items()}
                 for room, groups in SETTINGS["branch_groups"].items() if room in ROOMS}

# Front-block wedge seats are never allotted at planning time. They are the
# transition pool: late admissions, medium transfers, seat disputes.
RESERVE_WEDGE = True

class _Palette(dict):
    """Block letter -> hex colour. A hall set up with blocks past G gets a
    colour from the cycle instead of a KeyError in the workbook or handout."""
    _CYCLE = ["C9721B", "12706B", "7A3560", "4F6B1E", "1F5D8C", "8A4B12", "5A3E8C",
              "2E7D6B", "9C3D3D", "3F5E9C", "7A6A1F", "4B4B4B"]

    def __missing__(self, key):
        k = str(key)
        return self._CYCLE[(ord(k[0]) - ord("A")) % len(self._CYCLE)] if k else "444444"

    def get(self, key, default=None):
        return self[key]


BLOCK_COLOUR = _Palette({
    "A": "C9721B", "B": "12706B", "C": "7A3560", "D": "4F6B1E",
    "E": "1F5D8C", "F": "8A4B12", "G": "5A3E8C",
})

# ---- signature sheet geometry (millimetres, A4 portrait) --------------------
PAGE_W, PAGE_H = 210.0, 297.0
# The generous top margin is a holdover from the automatic scanner that used
# to read these sheets: it kept the table clear of a 25 mm corner marker band.
# The markers are gone, but the margin is kept so newly printed sheets line up
# page-for-page with every set already printed and filed.
MARGIN_X, MARGIN_TOP, MARGIN_BOT = 14.0, 46.0, 28.0
ROW_H = 8.4
COL_SEAT, COL_ROLL, COL_NAME = 20.0, 27.0, 62.0
COL_SIGN = PAGE_W - 2 * MARGIN_X - COL_SEAT - COL_ROLL - COL_NAME
ROWS_PER_PAGE = int((PAGE_H - MARGIN_TOP - MARGIN_BOT) // ROW_H)
