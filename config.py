"""Everything course- and room-specific lives here.

To run the toolkit for your own course, edit the COURSE block just below
(code, title, session, institute, mail domain, file names) and the
instructors further down; for your own halls, edit ROOMS and BRANCH_GROUPS.
No other file names the course.
"""
from dataclasses import dataclass, field

# ── The course ──────────────────────────────────────────────────────────────
COURSE_CODE = "DEMO101"
COURSE_TITLE = "Demo Course"
COURSE = f"{COURSE_CODE} {COURSE_TITLE}"         # printed on every sheet, page and mail
SESSION = "AY 2026-27 Sem 1"

INSTITUTE = "Example Institute of Technology"
INSTITUTE_SHORT = "EIT"
DEPARTMENT = "Department of Physics"
# Student and TA addresses are derived, never stored: <roll>@EMAIL_DOMAIN.
EMAIL_DOMAIN = "example.edu"
# Where the published "find your block / seat" page lives, if anywhere.
LOOKUP_URL = ""

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


LHC110 = [
    Block("A", "Front left",      6,  6, (),               "front"),
    Block("B", "Front centre-L",  7,  6, (3, 2, 2, 1),     "front"),
    Block("C", "Front centre-R",  7,  6, (4, 3, 2, 2, 1),  "front"),
    Block("D", "Front right",     6,  6, (),               "front"),
    Block("E", "Rear left",       14, 10, (),              "rear"),
    Block("F", "Rear centre",     13, 10, (),              "rear"),
    Block("G", "Rear right",      14, 10, (),              "rear"),
]

LHC2_101 = [
    Block("A", "Left",         11, 4, (), "front"),
    Block("B", "Centre-Left",  11, 4, (), "front"),
    Block("C", "Centre-Right", 11, 4, (), "front"),
    Block("D", "Right",        11, 4, (), "front"),
]

ROOMS = {"LHC110": LHC110, "LHC2 101": LHC2_101}

# Who the course belongs to. Printed on anything a student or an invigilator
# reads, and the source of the Cc list on every mail the toolkit sends, so the
# names and the addresses cannot drift apart.
COURSE_INSTRUCTORS = [
    ("Dr. Instructor One", "instructor.one@example.edu"),
    ("Dr. Instructor Two", "instructor.two@example.edu"),
]


def instructor_names(sep: str = ", ") -> str:
    return sep.join(name for name, _ in COURSE_INSTRUCTORS)


def instructor_emails(sep: str = ",") -> str:
    return sep.join(email for _, email in COURSE_INSTRUCTORS)

# Optional branch-grouped seating: which
# 2-letter department codes (see seating.py's `branch_code`) sit together in
# which block, in room order. A block with no entry here just isn't part of
# the grouping - it's filled by the catch-all pass in `allocate_grouped`
# (backlog students, and any branch not listed for any block). Only current
# B26 branches are grouped; backlog students are never split by branch, per
# the arrangement sheet ("Backlogs - all branches" as one lump total).
BRANCH_GROUPS = {
    "LHC110": {
        "A": ["BB", "CI", "MA"],
        "B": ["PH", "MT", "ME"],
        "C": ["CM", "EE", "CY"],
        "D": ["AE", "CS"],
        "F": ["CH", "EC"],
    },
    "LHC2 101": {
        "A": ["CS", "ME"],
        "B": ["EC", "AE", "CH"],
        "C": ["CI", "EE", "BB"],
        "D": ["CM", "MA", "PH", "CY"],
        # MT ("Overflow" in the arrangement sheet) isn't assigned a block -
        # the sheet's dedicated overflow bucket doesn't exist in this room's
        # current 4x44 layout, so those students fall through to catch-all.
    },
}

# Front-block wedge seats are never allotted at planning time. They are the
# transition pool: late admissions, medium transfers, seat disputes.
RESERVE_WEDGE = True

BLOCK_COLOUR = {
    "A": "C9721B", "B": "12706B", "C": "7A3560", "D": "4F6B1E",
    "E": "1F5D8C", "F": "8A4B12", "G": "5A3E8C",
}

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
