#!/usr/bin/env python3
"""Which TA invigilates which block, and how many are needed.

The rule the course uses, in its own words: LHC 110's front blocks take one TA
each; its long back blocks take two each; a smaller room takes two for the
room. That is a statement about sight-lines, not about student counts - a
front block is six seats wide and visible end to end from the aisle, while
E/F/G run fourteen rows deep and one person standing at the front cannot see
the back row's desks.

`requirements()` turns an allocation into that list of posts;
`auto_assign()` fills them from the TA duty sheet, Evaluation-duty TAs first
(exam duty is their assignment; the Attendance TAs are the weekly-class
roster), heads last since they supervise rather than stand. Deterministic, so
the same exam always produces the same sheet - a TA who checks twice sees the
same answer.
"""
from __future__ import annotations

import pandas as pd

import config as C

# Per-block TA counts for halls this project knows the shape of, keyed by the
# `Block.section` in `config.ROOMS`.
SECTION_TAS = {"front": 1, "rear": 2}
# A room with no known geometry (LHC 206/207/304/306/307 in the exam sheet)
# gets TAs for the room rather than per block: they are single small rooms
# whose blocks are all in one line of sight.
PER_ROOM_TAS = 2
WHOLE_ROOM = "-"          # the "block" value used for a room-level post

# Attendance TAs first. They are the ones who stand in these blocks every
# week for the class register, so they already know the room, the block
# boundaries and the faces - an invigilator who has to work out where block F
# ends is an invigilator not watching anyone. Evaluation TAs fill what's left,
# heads last: they float rather than hold a post.
DUTY_ORDER = ["Attendance", "Evaluation", "Evaluation Head", "Attendance Head"]

# Jobs that belong to the exam but not to a block. Question papers are
# collected and printed before the hall opens, so this is a real posting with
# a name against it - "someone will do it" is how a hall ends up with 364
# students and 340 papers.
SUPPORT_ROLES = {"Question paper printing": 2}
HEAD_DUTIES = {"Evaluation Head", "Attendance Head"}

# The heads hold no block - they move between rooms - but "no block" is not
# "no responsibility", and a pack that never names them reads as though the
# exam runs itself. Their scope comes from the duty sheet's own Notes column
# ("Minor and Major Exam Management", "Attendance, Seating and Quiz
# Management"); these are the concrete jobs that scope implies on the day.
HEAD_RESPONSIBILITIES = {
    "Evaluation Head": [
        "Question paper custody and distribution to rooms",
        "Invigilator briefing and room readiness",
        "Script collection, count and handover",
        "Decisions on any irregularity during the exam",
    ],
    "Attendance Head": [
        "Seating lists, block sheets and roll verification",
        "Absentee list after the exam",
    ],
}
# The Evaluation Head is overall in charge of every exam, quizzes included.
# The duty sheet's wording splits them ("... and Quiz Management" against the
# Attendance Head), but the course settled on one person answering for the
# exam whatever it is called, so there is never a question of whose call an
# irregularity is. He also briefs the invigilators. The Exam seating page can
# still switch the lead for a single exam.
IN_CHARGE_BY_EXAM = {"quiz": "Evaluation Head", "minor": "Evaluation Head",
                     "major": "Evaluation Head"}


def in_charge_duty(exam_title: str) -> str:
    t = (exam_title or "").strip().lower()
    for key, duty in IN_CHARGE_BY_EXAM.items():
        if t.startswith(key):
            return duty
    return "Evaluation Head"


def head_duties(tas: pd.DataFrame, exam_title: str = "") -> pd.DataFrame:
    """One row per head responsibility, with the head named and the overall
    in-charge marked. Empty frame if the duty sheet lists no heads - better a
    blank section than an invented one."""
    if not len(tas):
        return pd.DataFrame(columns=["Duty", "TA", "Roll", "Responsibility", "InCharge"])
    lead = in_charge_duty(exam_title)
    rows = []
    for duty in ("Evaluation Head", "Attendance Head"):
        who = tas[tas.Duty == duty]
        if who.empty:
            continue
        ta = who.iloc[0]
        for i, job in enumerate(HEAD_RESPONSIBILITIES.get(duty, [])):
            rows.append({"Duty": duty, "TA": ta.Name, "Roll": ta.Roll,
                         "Responsibility": job, "InCharge": duty == lead and i == 0})
    return pd.DataFrame(rows, columns=["Duty", "TA", "Roll", "Responsibility", "InCharge"])


def block_section(room_label: str, block_key: str) -> str | None:
    """`front` / `rear` from `config.ROOMS`, or None for a room this project
    only knows through the exam sheet."""
    room = C.ROOMS.get(room_label.replace(" ", ""))
    if not room:
        return None
    return next((b.section for b in room if b.key == block_key), None)


def requirements(alloc: pd.DataFrame, section_tas: dict | None = None,
                 per_room: int | None = None) -> pd.DataFrame:
    """One row per post to be filled: Room, Block, TAs needed, students there.

    `section_tas` / `per_room` override the defaults for one exam - three per
    long block instead of two is a reasonable call for a full hall, and the
    duty sheet has the people for it. Passing them keeps that decision with
    the exam rather than editing the module.

    Only blocks that actually hold students are counted - an empty block needs
    no one watching it.
    """
    section_tas = {**SECTION_TAS, **(section_tas or {})}
    per_room = PER_ROOM_TAS if per_room is None else int(per_room)
    rows = []
    for room in dict.fromkeys(alloc.Room):
        sub = alloc[alloc.Room == room]
        known = C.ROOMS.get(room.replace(" ", ""))
        if known:
            for block in dict.fromkeys(sub.Block):
                n = int((sub.Block == block).sum())
                if not n:
                    continue
                section = block_section(room, block) or "front"
                rows.append({"Room": room, "Block": block,
                             "Section": section, "Students": n,
                             "TAs": section_tas.get(section, 1)})
        else:
            rows.append({"Room": room, "Block": WHOLE_ROOM, "Section": "room",
                         "Students": int(len(sub)), "TAs": per_room})
    return pd.DataFrame(rows, columns=["Room", "Block", "Section", "Students", "TAs"])


def ta_pool(tas: pd.DataFrame) -> pd.DataFrame:
    """The TA sheet ordered the way posts should be filled."""
    df = tas.copy()
    df["_k"] = [DUTY_ORDER.index(d) if d in DUTY_ORDER else len(DUTY_ORDER) for d in df.Duty]
    return df.sort_values(["_k", "Name"]).drop(columns="_k").reset_index(drop=True)


def auto_assign(reqs: pd.DataFrame, tas: pd.DataFrame,
                exclude: list | None = None) -> pd.DataFrame:
    """Fill every post, in room order, from the ordered pool.

    Runs out rather than doubling anyone up: a TA in two places at once is not
    an assignment, it's a gap that looks filled. Unfilled posts come back with
    a blank name so the shortfall is visible on the sheet and on screen.

    `exclude` holds people already given an off-block job (printing). They are
    kept out of the block pool, since a TA at the printer cannot also be
    standing in a block - but only while the remaining pool still covers every
    post. Better a double-booking the sheet shows than a post left empty
    because someone was reserved for a job that takes twenty minutes.
    """
    pool = ta_pool(tas)
    if exclude:
        trimmed = pool[~pool.Name.isin(set(exclude))]
        if len(trimmed) >= int(reqs.TAs.sum()):
            pool = trimmed
    out, i = [], 0
    for r in reqs.itertuples():
        for _ in range(int(r.TAs)):
            if i < len(pool):
                ta = pool.iloc[i]
                out.append({"Room": r.Room, "Block": r.Block, "Section": r.Section,
                            "Students": r.Students, "TA": ta.Name, "Roll": ta.Roll,
                            "Duty": ta.Duty})
                i += 1
            else:
                out.append({"Room": r.Room, "Block": r.Block, "Section": r.Section,
                            "Students": r.Students, "TA": "", "Roll": "", "Duty": ""})
    return pd.DataFrame(out, columns=["Room", "Block", "Section", "Students", "TA", "Roll", "Duty"])


def assign_support(tas: pd.DataFrame, blocks: pd.DataFrame | None = None,
                   roles: dict | None = None, names: list | None = None) -> pd.DataFrame:
    """Fill the non-block duties.

    `names` posts specific people and is taken as given - **including a head**.
    The automatic pick avoids heads (they supervise, and the course asked for
    these jobs to sit elsewhere), but that is a default, not a rule the course
    is bound by: naming someone is an instruction, not a suggestion.

    Automatically, TAs not posted to a block are chosen first: printing
    happens before the hall opens, so a posted TA *could* do it, but sending
    someone who has to be standing in block E at the same time is how a job
    silently doesn't happen.
    """
    roles = roles or SUPPORT_ROLES
    posted_all = set(blocks.TA) if blocks is not None and len(blocks) else set()

    if names:
        rows = []
        by_name = {r.Name: r for r in tas.itertuples()} if len(tas) else {}
        for role, n in roles.items():
            for who in list(names)[:int(n)] if n else []:
                ta = by_name.get(who)
                rows.append({"Role": role, "TA": who,
                             "Roll": getattr(ta, "Roll", ""), "Duty": getattr(ta, "Duty", ""),
                             "AlsoInvigilating": who in posted_all})
            for _ in range(max(0, int(n) - len(list(names)))):
                rows.append({"Role": role, "TA": "", "Roll": "", "Duty": "",
                             "AlsoInvigilating": False})
        return pd.DataFrame(rows, columns=["Role", "TA", "Roll", "Duty", "AlsoInvigilating"])

    pool = ta_pool(tas)
    pool = pool[~pool.Duty.isin(HEAD_DUTIES)]
    posted = posted_all
    free = pool[~pool.Name.isin(posted)]
    order = pd.concat([free, pool[pool.Name.isin(posted)]]).reset_index(drop=True)

    rows, i = [], 0
    for role, n in roles.items():
        for _ in range(int(n)):
            if i < len(order):
                ta = order.iloc[i]
                rows.append({"Role": role, "TA": ta.Name, "Roll": ta.Roll, "Duty": ta.Duty,
                             "AlsoInvigilating": ta.Name in posted})
                i += 1
            else:
                rows.append({"Role": role, "TA": "", "Roll": "", "Duty": "",
                             "AlsoInvigilating": False})
    return pd.DataFrame(rows, columns=["Role", "TA", "Roll", "Duty", "AlsoInvigilating"])


def by_block(assignment: pd.DataFrame) -> dict[tuple[str, str], list[str]]:
    """{(room, block): [TA names]} - what the PDF and the seat map print."""
    out: dict[tuple[str, str], list[str]] = {}
    for r in assignment.itertuples():
        if r.TA:
            out.setdefault((r.Room, r.Block), []).append(r.TA)
    return out


def shortfall(assignment: pd.DataFrame) -> int:
    return int((assignment.TA == "").sum())


def summary(reqs: pd.DataFrame, tas: pd.DataFrame, roles: dict | None = None) -> str:
    roles = SUPPORT_ROLES if roles is None else roles
    need = int(reqs.TAs.sum()) + sum(int(v) for v in roles.values())
    have = len(tas)
    heads = int(tas.Duty.isin(HEAD_DUTIES).sum()) if len(tas) else 0
    verb = "enough" if have >= need else f"**{need - have} short**"
    extra = f" (incl. {sum(int(v) for v in roles.values())} off-block)" if roles else ""
    return (f"{need} TA posts to fill{extra} · {have} TAs on the duty sheet"
            + (f", {heads} of them heads" if heads else "") + f" - {verb}.")
