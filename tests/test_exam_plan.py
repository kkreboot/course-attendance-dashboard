"""The exam side: persistence, geometry, mixing, and what the published page
is allowed to contain."""
import json
import re
from pathlib import Path

import pandas as pd
import pytest

import config as C
import exam_plan
import finder
import posters


@pytest.fixture
def alloc():
    return pd.DataFrame({
        "Roll": ["B26CS1501", "B26CS1902", "B26ME1501", "B26ME1902"],
        "Name": ["A One", "B Two", "C Three", "D Four"],
        "Room": ["LHC 308"] * 4,
        "Block": ["A", "A", "B", "B"],
        "Seat": ["1", "3", "1", "3"],
    })


def test_save_and_load_round_trip(alloc, tmp_path):
    exam_plan.save_exam(alloc, title="Minor Exam", when="16 Sep", venue="LHC 308",
                        root=tmp_path)
    listed = exam_plan.list_exams(root=tmp_path)
    assert [m["title"] for m in listed] == ["Minor Exam"]
    assert listed[0]["students"] == 4 and listed[0]["rooms"] == ["LHC 308"]
    back, meta = exam_plan.load_exam("minor-exam", root=tmp_path)
    assert list(back.Roll) == list(alloc.Roll) and meta["when"] == "16 Sep"


def test_reallocating_replaces_rather_than_duplicates(alloc, tmp_path):
    exam_plan.save_exam(alloc, title="Minor Exam", root=tmp_path)
    exam_plan.save_exam(alloc.head(2), title="Minor Exam", root=tmp_path)
    listed = exam_plan.list_exams(root=tmp_path)
    assert len(listed) == 1 and listed[0]["students"] == 2


def test_interleave_separates_branches():
    rolls = [f"B26CS10{i:02d}" for i in range(10)] + [f"B26ME10{i:02d}" for i in range(10)]
    mixed = exam_plan.interleave_by_branch(rolls)
    assert sorted(mixed) == sorted(rolls), "a student was dropped or duplicated"
    adjacent_same = sum(exam_plan.branch_of(a) == exam_plan.branch_of(b)
                        for a, b in zip(mixed, mixed[1:]))
    assert adjacent_same == 0
    assert exam_plan.interleave_by_branch(rolls) == mixed, "not deterministic"


def test_student_page_carries_no_roll_numbers_or_names(alloc, tmp_path):
    out = tmp_path / "find.html"
    exam_plan.build_finder(alloc, str(out), meta={"title": "Minor Exam"})
    html = out.read_text(encoding="utf-8")
    for roll in alloc.Roll:
        assert roll not in html, "a roll number was published in the clear"
        assert finder.roll_key(roll) in html, "a student cannot look themselves up"
    for name in alloc.Name:
        assert name not in html


def test_seat_map_is_for_staff_and_does_carry_rolls(alloc, tmp_path):
    class _Block:
        def __init__(self, key, rows):
            self.key, self.rows = key, rows

    class _Room:
        def __init__(self):
            self.blocks = {"A": _Block("A", [["1", "2", "3"]]),
                           "B": _Block("B", [["1", "2", "3"]])}

    out = tmp_path / "map.html"
    exam_plan.build_seat_map(alloc, {"LHC 308": _Room()}, str(out),
                             meta={"title": "Minor Exam", "spaced": True})
    html = out.read_text(encoding="utf-8")
    assert "B26CS1501" in html
    payload = json.loads(re.search(r"const P=(\{.*?\});", html, re.S).group(1))
    grid = payload["rooms"]["LHC 308"]["A"]
    used = [c for row in grid for c in row if c["used"]]
    empty = [c for row in grid for c in row if not c["used"]]
    assert [c["seat"] for c in used] == ["1", "3"]
    assert empty, "unused seats must still be drawn - spacing is only legible with the gaps"


def test_scheduled_exams_are_defined():
    keys = [e["key"] for e in exam_plan.SCHEDULE]
    assert keys == ["quiz1", "minor", "major"], "schedule should read in date order"
    assert exam_plan.when_label(exam_plan.scheduled("quiz1")) == \
        "31 August 2026, Monday (08:00 PM, 30–45 minutes)"
    assert exam_plan.when_label(exam_plan.scheduled("minor")) == \
        "16 September 2026, Wednesday (04:30 PM to 06:30 PM)"
    assert exam_plan.when_label(exam_plan.scheduled("major")).startswith("21 November 2026, Saturday")
    # every exam is sat by both batches; nothing in the toolkit splits them
    assert all(e["batches"] == "English + Hindi" for e in exam_plan.SCHEDULE)

    from datetime import datetime
    for e in exam_plan.SCHEDULE:
        parsed = datetime.strptime(e["date"], "%d %B %Y")
        assert parsed.strftime("%A") == e["day"], f"{e['title']}: {e['date']} is not a {e['day']}"


def test_lhc110_is_shaped_like_the_hall_not_one_long_row():
    """The exam sheet gives LHC 110's blocks only as totals. Left flat, a
    block is one 140-seat line - which both draws wrong and makes "alternate
    seating" mean every second chair in a single row, seating people directly
    in front of and behind each other."""
    import config as C
    import exam_rooms

    rooms = exam_rooms.load_exam_rooms()
    room = rooms.get("LHC 110")
    if room is None:
        pytest.skip("LHC 110 not in this workbook")

    for key, block in room.blocks.items():
        assert len(block.rows) > 1, f"block {key} is still one flat row"
        assert max(len(r) for r in block.rows) <= 15

    geom = {b.key: b for b in C.ROOMS["LHC110"]}
    e = room.blocks["E"]
    assert len(e.rows) == geom["E"].rows and len(e.rows[0]) == geom["E"].cols
    assert sum(len(r) for r in e.rows) == geom["E"].total

    # alternate rows *and* alternate seats: 1,3,5,7,9 then the row after next
    assert e.spaced_seats[:6] == ["1", "3", "5", "7", "9", "21"]


def test_unknown_rooms_still_never_draw_one_endless_line(alloc, tmp_path):
    class _Block:
        def __init__(self, key, rows):
            self.key, self.rows = key, rows

        @property
        def dense_seats(self):
            return [s for r in self.rows for s in r]

    class _Room:
        def __init__(self):
            self.blocks = {"A": _Block("A", [[str(i) for i in range(1, 61)]]),
                           "B": _Block("B", [["1", "2", "3"]])}

    wide = pd.DataFrame({"Roll": ["B26CS1501"], "Name": ["A"], "Room": ["Somewhere Else"],
                         "Block": ["A"], "Seat": ["1"]})
    grid = exam_plan.room_layout(wide, {"Somewhere Else": _Room()})
    rows = grid["Somewhere Else"]["A"]
    assert len(rows) > 1 and max(len(r) for r in rows) <= 15


def test_three_spacing_methods():
    """Three ways to fill a hall, and the middle one is the point of this test:
    a gap beside each student, every row used."""
    import exam_rooms

    block = exam_rooms.ExamBlock(key="A", rows=[[str(r * 10 + c) for c in range(1, 11)]
                                                for r in range(4)])
    every = block.seats("every")
    side = block.seats("side")
    alternate = block.seats("alternate")

    assert len(every) == 40
    assert len(side) == 20 and len(alternate) == 10

    # "side": every row appears, alternate seats within it
    rows_used = {int(s) // 10 for s in side}
    assert rows_used == {0, 1, 2, 3}, "a row was skipped - that's the other mode"
    assert side[:5] == ["1", "3", "5", "7", "9"]
    assert "11" in side and "13" in side, "the row behind must be usable"

    # "alternate": rows 0 and 2 only
    assert {int(s) // 10 for s in alternate} == {0, 2}
    assert "11" not in alternate

    # nobody is beside anybody in either spaced mode
    for mode in ("side", "alternate"):
        picked = {int(s) for s in block.seats(mode)}
        assert not any(n + 1 in picked for n in picked if n % 10 != 0), \
            f"{mode} seated two students side by side"


def test_spacing_mode_accepts_the_old_boolean():
    """Exams saved before the third method existed stored `spaced: true`."""
    import exam_rooms

    block = exam_rooms.ExamBlock(key="A", rows=[["1", "2", "3", "4"], ["5", "6", "7", "8"]])
    assert block.seats(True) == block.seats("alternate")
    assert block.seats(False) == block.seats("every")
    with pytest.raises(ValueError, match="Unknown spacing"):
        block.seats("sideways")


def test_lhc110_transition_seats_are_never_allotted():
    """Blocks B and C taper into wedge seats behind the last full row. The
    classroom side holds those back as the transition pool; an exam must not
    seat anyone there either - but they still have to be *drawn*, or the gap
    reads as a missing student."""
    import exam_rooms

    rooms = exam_rooms.load_exam_rooms()
    room = rooms.get("LHC 110")
    if room is None:
        pytest.skip("LHC 110 not in this workbook")

    b, c = room.blocks["B"], room.blocks["C"]
    assert len(b.reserved) == 8 and len(c.reserved) == 12
    assert sum(len(r) for r in b.rows) == 50, "the wedge must still be drawn"
    assert len(b.dense_seats) == 42, "wedge seats leaked into the allottable list"

    for mode in ("every", "side", "alternate"):
        for block in (b, c):
            assert not (set(block.seats(mode)) & block.reserved), \
                f"{mode} allotted a transition seat"

    # blocks without a taper are untouched
    assert not room.blocks["A"].reserved and len(room.blocks["A"].dense_seats) == 36
    assert room.capacity("every") == 566


def test_reserved_seats_are_drawn_and_flagged_on_the_map(tmp_path):
    import exam_rooms

    block = exam_rooms.ExamBlock(key="B", rows=[["1", "2"], ["3", "4"]], reserved={"4"})

    class _Room:
        blocks = {"B": block}

    alloc = pd.DataFrame({"Roll": ["B26CS1501"], "Name": ["A"], "Room": ["R"],
                          "Block": ["B"], "Seat": ["1"]})
    grid = exam_plan.room_layout(alloc, {"R": _Room()})["R"]["B"]
    flat = {c["seat"]: c for row in grid for c in row}
    assert set(flat) == {"1", "2", "3", "4"}, "a reserved seat vanished from the drawing"
    assert flat["4"]["held"] and not flat["4"]["used"]
    assert flat["1"]["used"] and not flat["1"]["held"]


def test_overflow_seats_respect_spacing_and_stay_reserved():
    """The transition pool is a last resort, not a free-for-all: an overflow
    seat still obeys the spacing everyone else got."""
    import exam_rooms

    block = exam_rooms.ExamBlock(
        key="B",
        rows=[["1", "2", "3", "4"], ["5", "6", "7", "8"], ["9", "10", "11", "12"]],
        reserved={"9", "10", "11", "12"},
    )
    # normal seating never touches them
    for mode in ("every", "side", "alternate"):
        assert not (set(block.seats(mode)) & block.reserved)

    assert block.overflow_seats("every") == ["9", "10", "11", "12"]
    assert block.overflow_seats("side") == ["9", "11"], "overflow ignored the side gap"
    assert block.overflow_seats("alternate") == ["9", "11"], "wrong row for alternate mode"

    assert exam_rooms.OVERFLOW_LIMIT == 2


def test_lhc110_has_an_overflow_pool_in_every_mode():
    import exam_rooms

    rooms = exam_rooms.load_exam_rooms()
    room = rooms.get("LHC 110")
    if room is None:
        pytest.skip("LHC 110 not in this workbook")
    for mode in ("alternate", "side", "every"):
        pool = sum(len(b.overflow_seats(mode)) for b in room.blocks.values())
        assert pool >= exam_rooms.OVERFLOW_LIMIT, f"no overflow available in {mode}"


def test_ta_posts_follow_the_sight_line_rule():
    """One TA per front block, two per long back block, two per smaller room -
    a statement about what one person can see, not about headcount."""
    import ta_duty

    alloc = pd.DataFrame({
        "Roll": [f"B26XX{i:04d}" for i in range(9)],
        "Room": ["LHC 110"] * 7 + ["LHC 206"] * 2,
        "Block": ["A", "B", "C", "D", "E", "F", "G", "A", "B"],
        "Seat": [str(i) for i in range(9)],
    })
    reqs = ta_duty.requirements(alloc).set_index(["Room", "Block"])
    for block in "ABCD":
        assert reqs.at[("LHC 110", block), "TAs"] == 1
    for block in "EFG":
        assert reqs.at[("LHC 110", block), "TAs"] == 2
    # a room this project has no geometry for is posted as a whole room
    assert reqs.at[("LHC 206", ta_duty.WHOLE_ROOM), "TAs"] == 2
    assert len(reqs) == 8


def test_auto_assign_never_puts_a_ta_in_two_places():
    import ta_duty

    reqs = pd.DataFrame({"Room": ["R1", "R2"], "Block": ["A", "B"],
                         "Section": ["front", "rear"], "Students": [10, 10], "TAs": [1, 2]})
    tas = pd.DataFrame({"Roll": ["T1", "T2", "T3", "T4"],
                        "Name": ["Eva One", "Att Two", "Eva Three", "Head Four"],
                        "Duty": ["Evaluation", "Attendance", "Evaluation", "Evaluation Head"],
                        "Day": "", "Notes": ""})
    a = ta_duty.auto_assign(reqs, tas)
    assert len(a) == 3 and a.TA.nunique() == 3, "a TA was posted twice"
    # Attendance TAs first - they stand in these blocks every week for the
    # class register - then Evaluation, heads last
    assert list(a.Duty)[0] == "Attendance"
    # a head is only reached once everyone else is posted; three posts and
    # three non-heads means the head stays free to float
    assert "Evaluation Head" not in list(a.Duty)
    assert ta_duty.shortfall(a) == 0


def test_short_ta_sheet_leaves_visible_gaps_rather_than_doubling_up():
    import ta_duty

    reqs = pd.DataFrame({"Room": ["R1"], "Block": ["E"], "Section": ["rear"],
                         "Students": [70], "TAs": [2]})
    tas = pd.DataFrame({"Roll": ["T1"], "Name": ["Only One"], "Duty": ["Evaluation"],
                        "Day": "", "Notes": ""})
    a = ta_duty.auto_assign(reqs, tas)
    assert len(a) == 2 and ta_duty.shortfall(a) == 1
    assert ta_duty.by_block(a) == {("R1", "E"): ["Only One"]}


def test_exam_pack_pdf_has_cover_grids_and_signatures(tmp_path):
    import exam_rooms
    import posters
    from pypdf import PdfReader

    alloc = pd.read_csv("out/exams/minor-exam/allocation.csv", dtype={"Roll": str})
    duties = {("LHC 110", "A"): ["Someone Named"]}
    out = tmp_path / "pack.pdf"
    posters.exam_pack(alloc, str(out), room_map=exam_rooms.load_exam_rooms(),
                      exam="Minor Exam", when="16 September 2026", duties=duties)
    r = PdfReader(str(out))
    assert len(r.pages) > 10
    cover = r.pages[0].extract_text()
    assert "Minor Exam" in cover and "Someone Named" in cover
    assert "INVIGILATORS" in cover.upper()
    all_text = " ".join((p.extract_text() or "") for p in r.pages)
    assert "SIGNATURE" in all_text.upper()
    assert str(alloc.Roll.iloc[0]) in all_text


def test_ta_counts_are_overridable_per_exam():
    """Three per long block is a call to make for one exam, not a module edit."""
    import ta_duty

    alloc = pd.DataFrame({
        "Roll": [f"B26XX{i:04d}" for i in range(5)],
        "Room": ["LHC 110"] * 4 + ["LHC 206"],
        "Block": ["A", "B", "E", "F", "A"],
        "Seat": [str(i) for i in range(5)],
    })
    base = ta_duty.requirements(alloc)
    assert int(base.TAs.sum()) == 1 + 1 + 2 + 2 + 2

    bumped = ta_duty.requirements(alloc, section_tas={"rear": 3}, per_room=3)
    by = bumped.set_index(["Room", "Block"]).TAs
    assert by[("LHC 110", "E")] == 3 and by[("LHC 110", "F")] == 3
    assert by[("LHC 110", "A")] == 1, "front blocks should be untouched"
    assert by[("LHC 206", ta_duty.WHOLE_ROOM)] == 3
    # the module defaults must not have been mutated by the override
    assert ta_duty.SECTION_TAS == {"front": 1, "rear": 2}


def test_printing_duty_excludes_heads_and_prefers_unposted_tas():
    """Two TAs print the papers, and never the heads - they supervise."""
    import ta_duty

    tas = pd.DataFrame({
        "Roll": ["H1", "H2", "T1", "T2", "T3"],
        "Name": ["Head One", "Head Two", "Free One", "Free Two", "Posted Three"],
        "Duty": ["Evaluation Head", "Attendance Head", "Attendance", "Attendance",
                 "Evaluation"],
        "Day": "", "Notes": ""})
    blocks = pd.DataFrame({"Room": ["R"], "Block": ["A"], "Section": ["front"],
                           "Students": [10], "TA": ["Posted Three"], "Roll": ["T3"],
                           "Duty": ["Evaluation"]})
    sup = ta_duty.assign_support(tas, blocks)

    assert list(sup.Role) == ["Question paper printing"] * 2
    assert set(sup.TA) == {"Free One", "Free Two"}, "a posted TA or a head was picked"
    assert not sup.AlsoInvigilating.any()


def test_printing_duty_falls_back_to_posted_tas_but_says_so():
    import ta_duty

    tas = pd.DataFrame({"Roll": ["T1"], "Name": ["Only One"], "Duty": ["Evaluation"],
                        "Day": "", "Notes": ""})
    blocks = pd.DataFrame({"Room": ["R"], "Block": ["A"], "Section": ["front"],
                           "Students": [10], "TA": ["Only One"], "Roll": ["T1"],
                           "Duty": ["Evaluation"]})
    sup = ta_duty.assign_support(tas, blocks)
    assert sup.iloc[0].TA == "Only One" and bool(sup.iloc[0].AlsoInvigilating)
    assert sup.iloc[1].TA == "", "an empty post must stay visibly empty"


def test_summary_counts_off_block_posts():
    import ta_duty

    reqs = pd.DataFrame({"Room": ["R"], "Block": ["A"], "Section": ["front"],
                         "Students": [10], "TAs": [1]})
    tas = pd.DataFrame({"Roll": ["T1"], "Name": ["One"], "Duty": ["Evaluation"],
                        "Day": "", "Notes": ""})
    text = ta_duty.summary(reqs, tas)
    assert "3 TA posts" in text and "off-block" in text and "short" in text


def test_grid_never_truncates_a_roll_number_or_splits_a_block(tmp_path):
    """A seat map printing `B26CI…` defeats its own purpose, and flipping
    pages mid-block while standing in front of it defeats the rest."""
    import exam_rooms
    import posters
    from pypdf import PdfReader

    alloc = pd.read_csv("out/exams/quiz-1/allocation.csv", dtype={"Roll": str})
    out = tmp_path / "pack.pdf"
    posters.exam_pack(alloc, str(out), room_map=exam_rooms.load_exam_rooms(),
                      exam="Quiz 1", when="31 August 2026")
    pages = [(p.extract_text() or "") for p in PdfReader(str(out)).pages]
    grids = [t for t in pages if " - Block " in t and "signatures" not in t]
    assert grids, "no seat grids in the pack"

    for text in grids:
        assert "…" not in text, "a roll number was cut short in a cell"
        for word in text.split():
            if word.startswith("B2"):
                assert len(word) >= 8, f"truncated roll {word!r}"

    # a *grid* must fit one page; a signature list may legitimately run on,
    # since a 70-student block cannot be signed on a single sheet
    split_grids = [t for t in grids if "(cont.)" in t]
    assert not split_grids, "a seat grid was split across pages"


def test_a_flat_row_block_is_rewrapped_but_real_geometry_is_not():
    """LHC 206 arrives as one 15-seat line and may be re-wrapped for print;
    LHC 110's 10-wide rows are the hall's own shape and must survive."""
    import exam_rooms

    rooms = exam_rooms.load_exam_rooms()
    if "LHC 110" not in rooms:
        pytest.skip("workbook missing")
    e = rooms["LHC 110"].blocks["E"]
    assert [len(r) for r in exam_rooms._wrap_rows(e.rows, width=10)] == [10] * 14

    flat = [[str(i) for i in range(1, 16)]]
    assert [len(r) for r in exam_rooms._wrap_rows(flat, width=10)] == [10, 5]


def test_the_attendance_workbook_is_a_superset_of_the_roll_list():
    """The roll-list export is a snapshot and can miss a late admission -
    B22PH903 was in the register and the classroom seating but not the export,
    so an exam allocated from the export left him with no seat. Exams read the
    workbook instead; this guards the assumption behind that."""
    import seating

    if not Path(seating.ATTENDANCE_WORKBOOK).exists():
        pytest.skip("master workbook not present")
    register = set(seating.load_attendance_roster().Roll)

    seated = set()
    for cohort in ("english", "hindi"):
        p = Path("out") / cohort / "allocation.csv"
        if p.exists():
            seated |= set(pd.read_csv(p, dtype={"Roll": str}).Roll)
    assert seated <= register, "someone is seated in class but absent from the register"


def test_every_registered_student_has_an_exam_seat():
    """Whatever the roster source, a saved exam must cover everyone on the
    register - an unseated student turns up on the day with nowhere to sit."""
    import seating

    if not Path(seating.ATTENDANCE_WORKBOOK).exists():
        pytest.skip("master workbook not present")
    register = set(seating.load_attendance_roster().Roll)

    for meta in exam_plan.list_exams():
        alloc = pd.read_csv(Path(meta["dir"]) / exam_plan.ALLOC_FILE, dtype={"Roll": str})
        missing = register - set(alloc.Roll)
        assert not missing, f"{meta['title']}: no seat for {sorted(missing)[:5]}"


def test_signature_sheets_are_per_block_in_the_big_halls(tmp_path):
    """The invigilator standing in block E should hold block E's sheet only -
    a room-wide list starts with two pages that belong to someone else. Small
    rooms are the opposite case and are covered by the room-wise test below."""
    import exam_rooms
    import posters
    from pypdf import PdfReader

    alloc = pd.read_csv("out/exams/quiz-1/allocation.csv", dtype={"Roll": str})
    out = tmp_path / "pack.pdf"
    posters.exam_pack(alloc, str(out), room_map=exam_rooms.load_exam_rooms(),
                      exam="Quiz 1", when="31 August 2026")
    pages = [(p.extract_text() or "") for p in PdfReader(str(out)).pages]
    # the cover mentions signature sheets in prose; a real one has the heading
    sig_pages = [t for t in pages if any("· signatures" in l for l in t.split("\n"))]
    assert sig_pages

    # every signature page names exactly one block, and only that block's rolls
    seen_blocks = set()
    for text in sig_pages:
        header = next((l for l in text.split("\n") if "signatures" in l), "")
        if "Block" not in header:
            continue                      # a small room's room-wide sheet
        block = header.split("Block ")[1].split(" ")[0]
        room = header.split(" - ")[0].strip()
        seen_blocks.add((room, block))
        rolls_here = {w for w in text.split() if w.startswith("B2") and len(w) >= 8}
        theirs = set(alloc[(alloc.Room == room) & (alloc.Block == block)].Roll)
        assert rolls_here <= theirs, f"{room} {block} sheet lists another block's students"

    # one sheet per occupied block in the block-wise halls
    import posters as _p
    occupied = {(r.Room, str(r.Block)) for r in alloc.itertuples()
                if _p._is_block_wise(r.Room)}
    assert seen_blocks == occupied, f"missing sheets for {occupied - seen_blocks}"


def test_signature_sheet_has_a_collection_tally(tmp_path):
    import exam_rooms
    import posters
    from pypdf import PdfReader

    alloc = pd.read_csv("out/exams/quiz-1/allocation.csv", dtype={"Roll": str}).head(20)
    out = tmp_path / "pack.pdf"
    posters.exam_pack(alloc, str(out), room_map=exam_rooms.load_exam_rooms(),
                      exam="Quiz 1", when="31 August 2026")
    text = " ".join((p.extract_text() or "") for p in PdfReader(str(out)).pages)
    assert "scripts collected" in text and "absent" in text


def test_heads_get_named_responsibilities_and_the_right_one_leads():
    """Heads hold no block, but a pack that never names them reads as though
    the exam runs itself. Who leads comes from the duty sheet's own wording:
    quizzes are the Attendance Head's, Minor/Major the Evaluation Head's."""
    import ta_duty

    tas = pd.DataFrame({
        "Roll": ["H1", "H2", "T1"],
        "Name": ["Eval Head", "Att Head", "Plain TA"],
        "Duty": ["Evaluation Head", "Attendance Head", "Evaluation"],
        "Day": "", "Notes": ""})

    quiz = ta_duty.head_duties(tas, "Quiz 1")
    assert set(quiz.TA) == {"Eval Head", "Att Head"}, "a head was left off the pack"
    # one person answers for the exam whatever it is called, quizzes included
    assert quiz[quiz.InCharge].iloc[0].Duty == "Evaluation Head"

    minor = ta_duty.head_duties(tas, "Minor Exam")
    assert minor[minor.InCharge].iloc[0].Duty == "Evaluation Head"

    # the invigilator briefing belongs to the head who leads
    briefing = quiz[quiz.Responsibility.str.contains("briefing")]
    assert list(briefing.Duty) == ["Evaluation Head"]
    assert ta_duty.head_duties(tas, "Major Exam")[lambda d: d.InCharge].iloc[0].TA == "Eval Head"

    # every listed head has at least one concrete job, and exactly one lead
    for frame in (quiz, minor):
        assert (frame.groupby("TA").size() >= 1).all()
        assert int(frame.InCharge.sum()) == 1

    # heads still never take a block or the printing duty
    reqs = pd.DataFrame({"Room": ["R"], "Block": ["A"], "Section": ["front"],
                         "Students": [10], "TAs": [1]})
    assert "Head" not in ta_duty.auto_assign(reqs, tas).iloc[0].Duty
    assert not set(ta_duty.assign_support(tas).TA) & {"Eval Head", "Att Head"}


def test_no_heads_on_the_sheet_means_no_invented_section():
    import ta_duty

    tas = pd.DataFrame({"Roll": ["T1"], "Name": ["Plain"], "Duty": ["Evaluation"],
                        "Day": "", "Notes": ""})
    assert ta_duty.head_duties(tas, "Quiz 1").empty


def test_exam_pack_prints_the_management_section(tmp_path):
    import exam_rooms
    import posters
    import ta_duty
    from pypdf import PdfReader

    alloc = pd.read_csv("out/exams/quiz-1/allocation.csv", dtype={"Roll": str}).head(30)
    tas = pd.DataFrame({"Roll": ["H1", "H2"], "Name": ["Eval Head", "Att Head"],
                        "Duty": ["Evaluation Head", "Attendance Head"], "Day": "", "Notes": ""})
    out = tmp_path / "pack.pdf"
    posters.exam_pack(alloc, str(out), room_map=exam_rooms.load_exam_rooms(), exam="Quiz 1",
                      heads=ta_duty.head_duties(tas, "Quiz 1"))
    cover = PdfReader(str(out)).pages[0].extract_text()
    assert "Exam management" in cover
    assert "Eval Head" in cover and "OVERALL IN CHARGE" in cover
    assert "Script collection" in cover


def test_small_rooms_print_room_wise_and_big_halls_block_wise(tmp_path):
    """Five grids and five signature sheets for forty people in LHC 206 is five
    things to lose; LHC 110's blocks each have their own invigilator and need
    their own pages."""
    import exam_rooms
    import posters
    from pypdf import PdfReader

    alloc = pd.read_csv("out/exams/quiz-1/allocation.csv", dtype={"Roll": str})
    out = tmp_path / "pack.pdf"
    posters.exam_pack(alloc, str(out), room_map=exam_rooms.load_exam_rooms(),
                      exam="Quiz 1", when="31 August 2026")
    titles = []
    for page in PdfReader(str(out)).pages:
        for line in (page.extract_text() or "").split("\n"):
            if " - " in line and "Seating:" not in line:
                titles.append(line.strip())
                break

    assert any(t.startswith("LHC 110 - Block A") for t in titles)
    assert any(t.startswith("LHC 110 - Block A · signatures") for t in titles)

    # whichever rooms this exam actually uses, not a hardcoded pair: the venue
    # changes between exams
    room_wise = [r for r in dict.fromkeys(alloc.Room) if not posters._is_block_wise(r)]
    assert room_wise, "expected at least one room-wise room in this allocation"
    for room in room_wise:
        assert f"{room} - seating" in titles, f"{room} should print one room-wide grid"
        assert f"{room} - signatures" in titles
        assert not [t for t in titles if t.startswith(f"{room} - Block")], \
            f"{room} was still split block by block"

    # the room-wide page still labels each group of seats - "Block C" in a
    # lettered room, "Row 3" in one numbered by row and column - and lists
    # every student
    text = " ".join((p.extract_text() or "") for p in PdfReader(str(out)).pages)
    first = room_wise[0]
    for block in sorted(set(alloc[alloc.Room == first].Block)):
        assert exam_rooms.group_label(first, block) in text
    for roll in alloc[alloc.Room == first].Roll:
        assert roll in text


def test_block_order_is_alphabetical():
    """The exam sheet hands the small rooms back E..A, which reads as though
    the room starts at the far end."""
    import exam_rooms
    import posters
    from pypdf import PdfReader

    alloc = pd.read_csv("out/exams/quiz-1/allocation.csv", dtype={"Roll": str})
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "pack.pdf"
        posters.exam_pack(alloc, str(out), room_map=exam_rooms.load_exam_rooms(),
                          exam="Quiz 1")
        blocks = []
        for page in PdfReader(str(out)).pages:
            for line in (page.extract_text() or "").split("\n"):
                if line.startswith("LHC 110 - Block") and "signatures" not in line:
                    blocks.append(line.split("Block ")[1].strip())
        assert blocks == sorted(blocks), f"grids out of order: {blocks}"


def test_exam_roster_is_roll_list_order_with_missing_students_first():
    """The roll list is the order the department reads students in, so an exam
    that follows it is checkable against their own paperwork. A student the
    export has missed goes at the *front*, not dropped and not buried in the
    last row of the last block."""
    import seating

    roll_file = C.ROLL_LIST
    if not Path(roll_file).exists() or not Path(seating.ATTENDANCE_WORKBOOK).exists():
        pytest.skip("source files not present")

    roster = exam_plan.exam_roster(roll_file)
    export = list(seating.load_roster(roll_file).Roll.astype(str).str.strip())
    register = set(seating.load_attendance_roster().Roll)

    missing = [r for r in register if r not in export]
    assert list(roster.Roll)[:len(missing)] == sorted(missing) or set(
        list(roster.Roll)[:len(missing)]) == set(missing)
    assert list(roster.Roll)[len(missing):] == export, "the export's order was not kept"
    assert set(roster.Roll) >= register and not roster.Roll.duplicated().any()


def test_saved_exams_seat_students_in_roster_order():
    for meta in exam_plan.list_exams():
        alloc = pd.read_csv(Path(meta["dir"]) / exam_plan.ALLOC_FILE, dtype={"Roll": str})
        if alloc.empty:
            continue
        assert alloc.iloc[0].Roll == "B22PH903", \
            f"{meta['title']}: the late admission should be seated first"


def test_group_mail_puts_students_where_asked_and_never_names_a_seat():
    import mailer

    rolls = ["B26CS1501", "B26ME1902", "B22PH903"]
    m = mailer.build_exam_block_mail(
        rolls, exam="Quiz 1", when="31 August 2026", room="LHC 110", block="E",
        cc="instructor.one@example.edu", bcc="p23ph0911@example.edu")

    assert m.to.count("@") == 3 and "b22ph903@example.edu" in m.to
    assert m.cc == "instructor.one@example.edu"
    assert m.bcc == "p23ph0911@example.edu", "a TA address leaked out of Bcc"
    assert "Block E" in m.body and mailer.SEAT_TOKEN_RE.search(m.body) is None

    hidden = mailer.build_exam_block_mail(rolls, exam="Quiz 1", when="", room="LHC 206",
                                          students_in="bcc", bcc="ta@example.edu")
    assert hidden.to == "" and hidden.bcc.count("@") == 4
    with pytest.raises(ValueError):
        mailer.build_exam_block_mail([], exam="Quiz 1", when="", room="LHC 206")


def test_ta_addresses_follow_the_same_rule_as_students():
    import mailer

    assert mailer.ta_email("P24PH0999") == "p24ph0999@example.edu"
    assert mailer.ta_email("M25IQT995") == "m25iqt995@example.edu"


def test_named_support_tas_are_kept_out_of_the_block_pool():
    """Someone at the printer cannot also be standing in a block - but only
    while the rest of the pool still covers every post."""
    import ta_duty

    tas = pd.DataFrame({
        "Roll": ["A1", "A2", "E1", "H1"],
        "Name": ["Att One", "Att Two", "Eva One", "Head One"],
        "Duty": ["Attendance", "Attendance", "Evaluation", "Evaluation Head"],
        "Day": "", "Notes": ""})
    reqs = pd.DataFrame({"Room": ["R"], "Block": ["A"], "Section": ["front"],
                         "Students": [10], "TAs": [1]})

    posted = ta_duty.auto_assign(reqs, tas, exclude=["Att One"])
    assert posted.iloc[0].TA == "Att Two", "the printer was still given a block"

    # a named head is honoured even though the automatic pick avoids heads
    sup = ta_duty.assign_support(tas, posted, names=["Head One", "Eva One"])
    assert list(sup.TA) == ["Head One", "Eva One"]
    assert not sup.AlsoInvigilating.any()

    # excluding everyone would leave the post empty, so the exclusion is dropped
    tight = ta_duty.auto_assign(reqs, tas.head(1), exclude=["Att One"])
    assert tight.iloc[0].TA == "Att One", "a post was left empty to protect a reservation"


def test_instructor_names_are_one_list_shared_by_print_and_mail():
    """The names on the pack and the addresses on the Cc line are the same two
    people; a change should not have to be made in two places."""
    import config as C
    import mailer

    assert C.instructor_names() == "Dr. Instructor One, Dr. Instructor Two"
    assert mailer.DEFAULT_ATTENDANCE_CC == C.instructor_emails()
    assert C.instructor_emails().count("@") == len(C.COURSE_INSTRUCTORS)


def test_instructors_are_printed_on_the_pack(tmp_path):
    import config as C
    import exam_rooms
    import posters
    from pypdf import PdfReader

    alloc = pd.read_csv("out/exams/quiz-1/allocation.csv", dtype={"Roll": str}).head(30)
    out = tmp_path / "pack.pdf"
    posters.exam_pack(alloc, str(out), room_map=exam_rooms.load_exam_rooms(), exam="Quiz 1")
    pages = [(p.extract_text() or "") for p in PdfReader(str(out)).pages]
    assert C.instructor_names() in pages[0], "cover does not name the instructors"
    assert any(C.instructor_names() in p for p in pages[1:]), \
        "no inner page carries the instructors"


def test_grid_fills_the_page_and_stays_on_one_sheet(tmp_path):
    """A grid that uses a third of the page is a grid nobody can read from the
    front of a hall. Cells scale to the space; a block stays on one sheet."""
    import exam_rooms
    import posters
    from pypdf import PdfReader

    alloc = pd.read_csv("out/exams/quiz-1/allocation.csv", dtype={"Roll": str})
    out = tmp_path / "pack.pdf"
    posters.exam_pack(alloc, str(out), room_map=exam_rooms.load_exam_rooms(),
                      exam="Quiz 1", when="31 August 2026")
    pages = [(p.extract_text() or "") for p in PdfReader(str(out)).pages]
    titles = []
    for t in pages:
        line = next((l for l in t.split("\n") if " - " in l and "Seating:" not in l), "")
        # the cover's prose also contains " - "; a grid page's title names a
        # block or the room's own seating page
        if line and "signatures" not in line and (" - Block " in line
                                                  or line.endswith("- seating")):
            titles.append(line)

    assert titles, "no grid pages"
    assert not [t for t in titles if "(cont.)" in t], f"a grid spilled: {titles}"
    assert "FRONT OF HALL" not in " ".join(pages), "the front-of-hall line is back"
    # every occupied block still has exactly one grid page
    occupied = {(r.Room, str(r.Block)) for r in alloc.itertuples()}
    rooms_block_wise = {r for r, _ in occupied if posters._is_block_wise(r)}
    expected = len([1 for r, _ in occupied if r in rooms_block_wise]) \
        + len({r for r, _ in occupied} - rooms_block_wise)
    assert len(titles) == expected


def test_notice_board_map_sizes_paper_per_room(tmp_path):
    """A different object from the pack's A4 grid: the whole room on one sheet,
    read from arm's length rather than held. A3 for the big halls only - a
    forty-student room on A3 is harder to read, not easier."""
    import exam_rooms
    import posters
    from pypdf import PdfReader

    alloc = pd.read_csv("out/exams/quiz-1/allocation.csv", dtype={"Roll": str})
    out = tmp_path / "board.pdf"
    posters.notice_board_map(alloc, str(out), room_map=exam_rooms.load_exam_rooms(),
                             exam="Quiz 1", when="31 August 2026")
    r = PdfReader(str(out))
    rooms = list(dict.fromkeys(alloc.Room))
    assert len(r.pages) == len(rooms), "one page per room"

    import posters as _p
    for page, room in zip(r.pages, rooms):
        w = round(float(page.mediabox.width) / 72 * 25.4)
        h = round(float(page.mediabox.height) / 72 * 25.4)
        expected = (420, 297) if room in _p.A3_ROOMS else (210, 297)
        assert (w, h) == expected, f"{room} printed {w}x{h}mm, expected {expected}"

    # every student of a room appears on that room's page, in full
    for page, room in zip(r.pages, rooms):
        text = page.extract_text() or ""
        assert room in text
        for roll in alloc[alloc.Room == room].Roll:
            assert roll in text, f"{roll} missing from the {room} board"
        assert "..." not in text and "…" not in text


def test_notice_board_groups_blocks_the_way_the_hall_is():
    """LHC 110's rear gallery and front section are drawn as separate bands,
    because that is how someone standing at the door sees the room - and within
    a band the blocks run in the hall's own left-to-right order, not
    alphabetically, which is not how the room is arranged."""
    import posters

    grids = {k: [] for k in "ABCDEFG"}
    sections = posters._sections_for("LHC 110", grids)
    assert sections == [posters.LHC110_BLOCK_ORDER["rear"],
                        posters.LHC110_BLOCK_ORDER["front"]]
    assert sections == [["G", "F", "E"], ["D", "C", "B", "A"]]

    # a block the workbook has not named yet still appears rather than vanishing
    assert posters._ordered(["E", "G", "Z"], posters.LHC110_BLOCK_ORDER["rear"]) \
        == ["G", "E", "Z"]

    # a room this project has no geometry for is one band of blocks
    assert posters._sections_for("LHC 206", {k: [] for k in "ABCDE"}) == [["A", "B", "C", "D", "E"]]
    rows105 = [f"R{n:02d}" for n in range(1, 11)]
    assert posters._sections_for("LHC 105", {k: [] for k in rows105}) == [rows105]


def test_lhc105_replaces_the_two_small_rooms():
    """LHC 206 and LHC 207 merged: ten blocks of fifteen, 150 seats, so a row
    spaced alternately still seats eight and the pair's 80 places survive."""
    import exam_rooms

    rooms = exam_rooms.load_exam_rooms()
    assert "LHC 105" in rooms, "LHC 105 is not offered as a venue"
    room = rooms["LHC 105"]
    assert sorted(room.blocks) == [f"R{n:02d}" for n in range(1, 11)]
    assert room.capacity("every") == 150
    assert room.capacity("side") == 80 == rooms["LHC 206"].capacity("side") \
        + rooms["LHC 207"].capacity("side")

    # it is big enough for A3, and its blocks are one line of sight, so the
    # pack prints it room-wise like the rooms it replaces
    import posters
    assert "LHC 105" in posters.A3_ROOMS
    assert not posters._is_block_wise("LHC 105")


def test_saved_exams_use_lhc105_and_seat_everyone():
    import seating

    if not Path(seating.ATTENDANCE_WORKBOOK).exists():
        pytest.skip("register not present")
    register = set(seating.load_attendance_roster().Roll)
    for meta in exam_plan.list_exams():
        alloc = pd.read_csv(Path(meta["dir"]) / exam_plan.ALLOC_FILE, dtype={"Roll": str})
        rooms = set(alloc.Room)
        assert rooms == {"LHC 110", "LHC 105"}, f"{meta['title']} uses {rooms}"
        assert set(alloc.Roll) >= register


def test_superseded_rooms_stay_visible_with_a_reason():
    """LHC 206 and 207 are gone from the choices, but a person looking for
    them should learn where they went rather than assume the parser broke."""
    import exam_rooms

    rooms = exam_rooms.load_exam_rooms()
    for old in ("LHC 206", "LHC 207"):
        assert old in rooms, f"{old} vanished from the list entirely"
        assert not rooms[old].supported
        assert "LHC 105" in rooms[old].note

    offered = [l for l, r in rooms.items() if r.supported]
    assert "LHC 105" in offered
    assert offered == sorted(offered, key=exam_rooms._room_sort_key)
    assert offered[0] == "LHC 105", "rooms should read as numbers, not as text"


def test_block_mail_names_room_and_block_and_leaves_the_seat_out():
    """The mail's job is to get a student into the right room and block; the
    seat is on the door sheet. It can carry seats on request, and does not by
    default."""
    import mailer

    rolls = ["B26CS1501", "B26ME1902"]
    seats = {"B26CS1501": "A-1", "B26ME1902": "B-3"}
    plain = mailer.build_exam_block_mail(rolls, exam="Quiz 1", when="31 August 2026",
                                         room="LHC 105", block="A")
    assert "LHC 105, Block A" in plain.body
    assert "Your seat number:" not in plain.body
    assert mailer.SEAT_TOKEN_RE.search(plain.body) is None

    asked = mailer.build_exam_block_mail(rolls, exam="Quiz 1", when="", room="LHC 105",
                                         block="A", seats=seats)
    assert "B26CS1501  A-1" in asked.body and "B26ME1902  B-3" in asked.body


def test_every_student_gets_exactly_one_block_mail():
    """One mail per block in every room, so each student is told their own
    room and block once and only once."""
    import exam_rooms
    import mailer
    import posters

    for meta in exam_plan.list_exams():
        alloc = pd.read_csv(Path(meta["dir"]) / exam_plan.ALLOC_FILE, dtype={"Roll": str})
        groups, seen = [], []
        for room in dict.fromkeys(alloc.Room):
            sub_room = alloc[alloc.Room == room]
            if exam_rooms.is_row_col_room(room):
                # no blocks to name: one mail for the room, the row is on the
                # plan at the door
                groups.append((room, "", sub_room))
                seen += list(sub_room.Roll)
                continue
            for block in sorted(set(sub_room.Block.astype(str))):
                sub = sub_room[sub_room.Block.astype(str) == block]
                groups.append((room, block, sub))
                seen += list(sub.Roll)

        assert sorted(seen) == sorted(alloc.Roll), "a student was missed or mailed twice"
        for room, block, sub in groups:
            m = mailer.build_exam_block_mail(list(sub.Roll), exam=meta["title"],
                                             when=meta.get("when", ""), room=room,
                                             block=block)
            if block:
                assert f"{room}, Block {block}" in m.body
            else:
                assert str(room) in m.body and "Block" not in m.body
                assert "row and column" in m.body
            assert mailer.SEAT_TOKEN_RE.search(m.body) is None
        # printing may still be room-wise; that is about paper, not the mail
        assert any(not posters._is_block_wise(r) for r in alloc.Room) or True


def test_lhc105_is_read_as_rows_and_columns_not_blocks(tmp_path):
    """LHC 110 keeps blocks - a block there is a place an invigilator stands
    in. LHC 105 is one flat hall, so it is numbered the way the hall itself is:
    row down the side, column across the top, and no block letter anywhere a
    student can see."""
    import exam_rooms
    import mailer
    import posters
    from pypdf import PdfReader

    assert exam_rooms.is_row_col_room("LHC 105")
    assert not exam_rooms.is_row_col_room("LHC 110")
    assert exam_rooms.group_label("LHC 105", "R03") == "Row 3"
    assert exam_rooms.group_label("LHC 110", "C") == "Block C"
    assert exam_rooms.place_label("LHC 105", "R03", "7") == "LHC 105, Row 3, Column 7"

    alloc = pd.read_csv("out/exams/quiz-1/allocation.csv", dtype={"Roll": str})
    sub = alloc[alloc.Room == "LHC 105"]
    assert not sub.empty
    assert all(exam_rooms.row_number(b) is not None for b in sub.Block), \
        "LHC 105 still carries block letters"

    out = tmp_path / "pack.pdf"
    posters.exam_pack(alloc, str(out), room_map=exam_rooms.load_exam_rooms(),
                      exam="Quiz 1", when="31 August 2026")
    text = " ".join((p.extract_text() or "") for p in PdfReader(str(out)).pages)
    assert "LHC 105 - seating" in text and "Row 1" in text and "Row 10" in text
    assert "Block R" not in text, "a row was printed as a block"

    m = mailer.build_exam_block_mail(list(sub.Roll)[:5], exam="Quiz 1",
                                     when="31 August 2026", room="LHC 105")
    assert "LHC 105" in m.body and "Block" not in m.body
    assert "row and column" in m.body
    assert mailer.SEAT_TOKEN_RE.search(m.body) is None


def test_wide_room_prints_landscape_a4_so_it_reads_in_print(tmp_path):
    """Fifteen columns on a portrait A4 is an 11mm desk and a 5pt roll number.
    The same sheet turned sideways gives an 18mm desk, so a wide room prints
    landscape and every other page stays portrait."""
    import exam_rooms
    import posters
    from pypdf import PdfReader

    alloc = pd.read_csv("out/exams/quiz-1/allocation.csv", dtype={"Roll": str})
    out = tmp_path / "pack.pdf"
    posters.exam_pack(alloc, str(out), room_map=exam_rooms.load_exam_rooms(),
                      exam="Quiz 1", when="31 August 2026")

    def mm_size(page):
        return (round(float(page.mediabox.width) / 72 * 25.4),
                round(float(page.mediabox.height) / 72 * 25.4))

    landscape = []
    for page in PdfReader(str(out)).pages:
        w, h = mm_size(page)
        assert {w, h} == {210, 297}, "the pack left A4 entirely"
        if w > h:
            landscape.append((page.extract_text() or "").split("\n")[2])
    # sideways pages are the wide seat grid and the two-up signature lists,
    # and nothing else
    assert "LHC 105 - seating" in landscape, "the wide seat grid stayed portrait"
    assert all(t == "LHC 105 - seating" or "signatures" in t for t in landscape), \
        f"a page turned sideways that should not have: {landscape}"


def test_notice_board_can_print_a4_only(tmp_path):
    """A printer with no A3 is not helped by an A3 sheet scaled down. In A4
    mode a hall that needed A3 is split across A4 sheets at the same cell size,
    each saying which sheet it is."""
    import exam_rooms
    import posters
    from pypdf import PdfReader

    alloc = pd.read_csv("out/exams/quiz-1/allocation.csv", dtype={"Roll": str})
    rooms = exam_rooms.load_exam_rooms()

    a3 = posters.notice_board_map(alloc, str(tmp_path / "auto.pdf"), room_map=rooms,
                                  exam="Quiz 1", when="31 August 2026")
    sizes = {(round(float(p.mediabox.width) / 72 * 25.4),
              round(float(p.mediabox.height) / 72 * 25.4))
             for p in PdfReader(a3).pages}
    assert (420, 297) in sizes, "the default stopped using A3 for the big halls"

    out = posters.notice_board_map(alloc, str(tmp_path / "a4.pdf"), room_map=rooms,
                                   exam="Quiz 1", when="31 August 2026", paper="a4")
    pages = PdfReader(out).pages
    for page in pages:
        w = round(float(page.mediabox.width) / 72 * 25.4)
        h = round(float(page.mediabox.height) / 72 * 25.4)
        assert {w, h} == {210, 297}, "A4 mode printed something that is not A4"

    text = " ".join((p.extract_text() or "") for p in pages)
    assert "sheet 1 of" in text and "sheet 2 of" in text, "sheets are not numbered"
    # every student is still on a sheet somewhere
    for roll in alloc.Roll:
        assert roll in text

    with pytest.raises(ValueError):
        posters.notice_board_map(alloc, str(tmp_path / "x.pdf"), room_map=rooms,
                                 exam="Quiz 1", paper="A5")


def test_empty_desks_are_drawn_narrow_in_every_room(tmp_path):
    """A spaced hall is mostly empty chairs. They stay on the sheet - an
    invigilator has to see which desks should be empty - but they are drawn
    narrow, so the width goes to the roll numbers instead. This holds for
    LHC 110's blocks as well as LHC 105's rows."""
    import exam_rooms
    import posters

    assert 0 < posters.NARROW_COL < 1 and 0 < posters.FLAT_CELL <= 1

    rows = [["1", "2", "3", "4"]]
    seat_roll = {("A", "1"): "B26XX1001", ("A", "3"): "B26XX1002"}
    seated = posters._seated_cols(rows, "A", seat_roll)
    assert seated == [True, False, True, False]
    widths = posters._col_widths(seated, 20.0)
    assert widths == [20.0, 20.0 * posters.NARROW_COL] * 2
    # a spaced block is narrower than four full columns, which is the width
    # the seated columns get back
    assert sum(widths) < 4 * 20.0
    assert posters._col_weight([True, True]) == 2.0

    # and the whole pack still builds, one grid page per block, nothing cut
    alloc = pd.read_csv("out/exams/quiz-1/allocation.csv", dtype={"Roll": str})
    out = tmp_path / "pack.pdf"
    posters.exam_pack(alloc, str(out), room_map=exam_rooms.load_exam_rooms(),
                      exam="Quiz 1", when="31 August 2026")
    from pypdf import PdfReader
    text = " ".join((p.extract_text() or "") for p in PdfReader(str(out)).pages)
    assert "…" not in text, "a roll number was cut short"
    for roll in alloc.Roll:
        assert roll in text


def test_lhc105_signature_list_is_two_up_on_landscape(tmp_path):
    """Eighty names down a portrait sheet is four pages to keep in order. Two
    lists side by side on landscape A4 is two sheets, which is what actually
    gets handed to an invigilator."""
    import exam_rooms
    import posters
    from pypdf import PdfReader

    alloc = pd.read_csv("out/exams/quiz-1/allocation.csv", dtype={"Roll": str})
    out = tmp_path / "pack.pdf"
    posters.exam_pack(alloc, str(out), room_map=exam_rooms.load_exam_rooms(),
                      exam="Quiz 1", when="31 August 2026")

    sig_pages = []
    for page in PdfReader(str(out)).pages:
        lines = (page.extract_text() or "").split("\n")
        title = lines[2] if len(lines) > 2 else ""
        if title.startswith("LHC 105 - signatures"):
            sig_pages.append((page, " ".join(lines)))

    assert len(sig_pages) == 2, f"expected 2 signature sheets, got {len(sig_pages)}"
    for page, _ in sig_pages:
        w = round(float(page.mediabox.width) / 72 * 25.4)
        h = round(float(page.mediabox.height) / 72 * 25.4)
        assert (w, h) == (297, 210), "the signature sheet is not landscape A4"

    text = " ".join(t for _, t in sig_pages)
    for r in alloc[alloc.Room == "LHC 105"].itertuples():
        assert r.Roll in text, f"{r.Roll} is on no signature sheet"
    # a serial number, not the row and column: the seat is on the plan, and a
    # signature sheet is read down a list of people
    assert "S. NO" in text and "SIGNATURE" in text
    assert "ROW" not in text and "COL" not in text
    # the serial runs on from LHC 110 rather than restarting at 1
    serials = posters.serial_numbers(alloc)
    here = [serials[r] for r in alloc[alloc.Room == "LHC 105"].Roll]
    assert min(here) > 1, "LHC 105 restarted the numbering"
    assert str(min(here)) in text and str(max(here)) in text
    assert "80 students" in text


def test_signature_serial_runs_on_across_blocks_and_rooms():
    """The heads collect one pile of sheets at the end. A serial that restarts
    in every block makes "number 47" a question rather than an answer, so it
    runs 1..N over the whole exam, in the order the sheets print."""
    import posters

    alloc = pd.read_csv("out/exams/quiz-1/allocation.csv", dtype={"Roll": str})
    serials = posters.serial_numbers(alloc)

    assert len(serials) == len(alloc), "a student is missing a serial"
    assert sorted(serials.values()) == list(range(1, len(alloc) + 1)), \
        "the serials are not 1..N without gaps or repeats"

    # each room takes one unbroken stretch, in the order the rooms print
    lo_hi = []
    for room in dict.fromkeys(alloc.Room):
        got = sorted(serials[r] for r in alloc[alloc.Room == room].Roll)
        assert got == list(range(got[0], got[-1] + 1)), f"{room} is interleaved"
        lo_hi.append((got[0], got[-1]))
    assert lo_hi == sorted(lo_hi) and lo_hi[0][0] == 1
    for (_, end), (start, _) in zip(lo_hi, lo_hi[1:]):
        assert start == end + 1, "a room restarted or skipped numbers"

    # and inside a block the serial follows the seat order the sheet reads in
    first = alloc[alloc.Room == "LHC 110"]
    blk = first[first.Block == "A"].copy()
    blk["n"] = [serials[r] for r in blk.Roll]
    order = blk.sort_values("Seat", key=lambda c: c.map(int)).n.tolist()
    assert order == sorted(order), "the serial jumps about inside a block"


def test_long_blocks_get_two_up_signature_sheets(tmp_path):
    """LHC 110's back blocks hold 65-70 people, which is three portrait sheets
    for one invigilator to keep in order. They print two lists to a landscape
    sheet, two sheets a block; a front block of twenty stays on its one page."""
    import exam_rooms
    import posters
    from pypdf import PdfReader

    alloc = pd.read_csv("out/exams/quiz-1/allocation.csv", dtype={"Roll": str})
    out = tmp_path / "pack.pdf"
    posters.exam_pack(alloc, str(out), room_map=exam_rooms.load_exam_rooms(),
                      exam="Quiz 1", when="31 August 2026")

    sheets = {}
    for page in PdfReader(str(out)).pages:
        lines = (page.extract_text() or "").split("\n")
        title = lines[2] if len(lines) > 2 else ""
        if "signatures" not in title:
            continue
        key = title.split(" · ")[0]
        if title.startswith("LHC 105"):
            key = "LHC 105 - signatures"
        w = round(float(page.mediabox.width) / 72 * 25.4)
        h = round(float(page.mediabox.height) / 72 * 25.4)
        sheets.setdefault(key, []).append((w, h))

    for block in ("E", "F", "G"):
        pages = sheets[f"LHC 110 - Block {block}"]
        assert len(pages) == 2, f"block {block} took {len(pages)} sheets"
        assert all(p == (297, 210) for p in pages), f"block {block} is not landscape"

    for block in ("A", "B", "C", "D"):
        pages = sheets[f"LHC 110 - Block {block}"]
        assert pages == [(210, 297)], f"block {block} should stay one portrait sheet"

    assert len(sheets["LHC 105 - signatures"]) == 2


def test_where_to_report_is_one_page_of_roll_ranges(tmp_path):
    """The sheet a student reads: roll number to room and block, one page, no
    seats and no names. A block that is not one unbroken stretch of roll
    numbers has to say so rather than quietly round over it."""
    import posters
    from pypdf import PdfReader

    alloc = pd.read_csv("out/exams/quiz-1/allocation.csv", dtype={"Roll": str})
    rows = posters.roll_ranges(alloc)

    # every student is covered exactly once by the ranges
    assert sum(r["n"] for r in rows) == len(alloc)
    rolls = sorted(str(r) for r in alloc.Roll)
    covered = []
    for r in rows:
        for lo, hi in r["runs"]:
            covered += rolls[rolls.index(lo):rolls.index(hi) + 1]
    assert sorted(covered) == rolls, "the ranges do not cover everyone exactly once"

    # the two students the roster gained at the end are their own stretch
    block_b = next(r for r in rows if r["room"] == "LHC 110" and r["block"] == "B")
    assert len(block_b["runs"]) == 2, "block B's second stretch was rounded over"

    # a range never spans two branches: "B26AE1943 to B26BB1932" reads as one
    # sweep of numbers but joins AE to BB
    for r in rows:
        for lo, hi in r["runs"]:
            assert posters._branch_of(lo) == posters._branch_of(hi), \
                f"{lo} to {hi} crosses a branch"

    # LHC 105 is one line for the room, since there is no block to send anyone to
    room_105 = [r for r in rows if r["room"] == "LHC 105"]
    assert len(room_105) == 1 and room_105[0]["block"] == "" and room_105[0]["rowcol"]

    # the seniors' block is described in words, not as a range spanning batches
    block_a = next(r for r in rows if r["block"] == "A")
    assert "every roll number from B22" in block_a["note"]

    out = tmp_path / "summary.pdf"
    posters.report_summary(alloc, str(out), exam="Quiz 1", when="31 August 2026",
                           late_note="Report to the front of LHC 110.")
    pages = PdfReader(str(out)).pages
    assert len(pages) == 1, "the summary spilled onto a second page"
    text = pages[0].extract_text() or ""
    for room in ("LHC 110", "LHC 105"):
        assert room in text
    lo, hi = next(run for r in rows for run in r["runs"] if run[0] != run[1])
    assert f"{lo} to {hi}" in text
    names = set(alloc.Name.astype(str))
    assert not [n for n in names if n and n in text], "a student name reached the summary"
    assert "…" not in text, "a line was cut short"
