"""What must never appear in a student mail, and who must never be mailed twice."""
import pandas as pd
import pytest

import config as C
import mailer


def test_address_is_derived_from_the_roll_number():
    assert mailer.student_email("B26BB1901") == "b26bb1901@example.edu"
    assert mailer.student_email(" b26bb1901 ") == "b26bb1901@example.edu"
    for bad in ("", "  ", "B26/BB", "a b"):
        with pytest.raises(ValueError):
            mailer.student_email(bad)


def test_transition_mail_never_names_a_seat():
    m = mailer.build_transition_mail("B26BB1901", "C", name="X", old_block="A", room="LHC110")
    assert "Block C" in m.subject
    # the word "seating" is fine; a seat *id* is not
    assert mailer.SEAT_TOKEN_RE.search(m.body) is None
    assert mailer.SEAT_TOKEN_RE.search(m.subject) is None
    # a seat id smuggled in through a free-text field must be refused, not printed
    with pytest.raises(ValueError, match="seat id"):
        mailer.build_transition_mail("B26BB1901", "C", room="LHC110",
                                     effective="from seat A-004 onwards")


def test_exam_mail_may_name_a_seat_but_can_be_told_not_to():
    m = mailer.build_exam_mail("B26BB1901", room="LHC110", block="B", seat="B-012", exam="Quiz 1")
    assert "B-012" in m.body
    with pytest.raises(ValueError, match="seat id"):
        mailer.build_exam_mail("B26BB1901", room="LHC110", block="B", seat="B-012",
                               exam="Quiz 1", include_seat=False, note="your seat is B-012")


def test_attendance_mail_carries_cc_and_no_seat():
    m = mailer.build_attendance_mail("B26CS1920", name="A", present=8, absent=6, held=14,
                                     remaining=18, asof="25/09/2026")
    assert m.cc == mailer.DEFAULT_ATTENDANCE_CC
    assert "Block" not in m.body and mailer.SEAT_TOKEN_RE.search(m.body) is None
    assert "present for 8 classes and absent for 6" in m.body
    assert "As of 25/09/2026 (14 classes held)" in m.body


def test_attendance_arithmetic_decides_which_template():
    """Every class missed costs three to make up at 75%, and once the remaining
    classes cannot cover that the notice has to stop saying "you can still make
    it" - the student would find out it was untrue."""
    # 8 of 14: 3*6 - 8 = 10 classes in a row gets back to 75%
    assert mailer.classes_to_recover(8, 6) == 10
    assert mailer.classes_to_recover(21, 0) == 0            # already above
    assert mailer.classes_to_recover(0, 1) == 3

    assert mailer.best_possible_percent(8, 6, 18) == pytest.approx(81.25)
    assert mailer.can_still_reach(8, 6, 18)
    assert not mailer.can_still_reach(4, 17, 18)            # 56.4% at best

    good = mailer.build_attendance_mail("B26CS1920", name="A", present=8, absent=6,
                                        held=14, remaining=18, asof="25/09/2026")
    assert good.subject == f"{C.COURSE_CODE} attendance below 75%: action needed"
    assert "attend the next 10 classes in a row" in good.body
    assert "18 classes remain in the course" in good.body
    assert "Excused absences are not counted" in good.body
    assert good.kind == "attendance_shortfall"

    lost = mailer.build_attendance_mail("B26CY1906", name="B", present=4, absent=17,
                                        held=21, remaining=18, asof="25/09/2026",
                                        meet_by="3 October 2026")
    assert lost.subject == f"{C.COURSE_CODE} attendance shortfall: please meet the course instructor"
    assert "the highest you can reach is 56.4%" in lost.body
    assert "Meet the course instructors by 3 October 2026" in lost.body
    assert "Submit supporting documents by 3 October 2026" in lost.body
    assert "in a row" not in lost.body, "the recoverable wording leaked into the second one"
    assert lost.kind == "attendance_unreachable"
    assert lost.meta["reachable"] is False


def test_attendance_mail_optional_lines_are_off_by_default():
    """A consequence line states policy, so it is never invented: blank means
    nothing is said. The meet-by date degrades to plain words rather than
    printing an empty placeholder."""
    m = mailer.build_attendance_mail("B26CY1906", name="B", present=4, absent=17,
                                     held=21, remaining=18)
    assert "Meet the course instructors as soon as possible" in m.body
    assert "{" not in m.body and "}" not in m.body
    assert "policy" not in m.body

    with_line = mailer.build_attendance_mail(
        "B26CY1906", name="B", present=4, absent=17, held=21, remaining=18,
        consequence="Students below the requirement may be barred from the exam.")
    assert "may be barred from the exam." in with_line.body


def test_attendance_mail_signs_off_as_the_ta():
    m = mailer.build_attendance_mail("B26CS1920", name="A", present=8, absent=6, held=14,
                                     remaining=18)
    assert m.body.rstrip().endswith(mailer.ATTENDANCE_SIGNATURE)
    assert "Regards," in m.body and "Best Regards," not in m.body


def test_escalation_levels():
    assert mailer.attendance_level(4) is None
    assert mailer.attendance_level(5) == "L1"
    assert mailer.attendance_level(12) == "L2"
    assert mailer.attendance_level(99) == "L3"


def test_unmailed_attendance_tracks_levels(tmp_path):
    log = tmp_path / "log.csv"
    students = pd.DataFrame({"Roll": ["A", "B"], "Absent": [6, 10], "Held": [14, 14],
                             "Present": [8, 4], "Percent": [57.0, 28.0], "Name": ["a", "b"]})
    first = mailer.unmailed_attendance(students, path=log)
    assert dict(zip(first.Roll, first.Level)) == {"A": "L1", "B": "L2"}

    mailer.log_sends([dict(timestamp="2026-01-01T00:00:00", kind="attendance_shortfall_L1",
                           roll="A", to="a@x", subject="s", status="sent", detail="")], log)
    assert list(mailer.unmailed_attendance(students, path=log).Roll) == ["B"]

    worse = students.copy()
    worse.loc[worse.Roll == "A", "Absent"] = 10          # A slips to the next level
    assert "A" in set(mailer.unmailed_attendance(worse, path=log).Roll)


def test_unmailed_transitions_compares_timestamps(tmp_path):
    log = tmp_path / "log.csv"
    trans = pd.DataFrame([{"timestamp": "2026-05-02T10:00:00", "roll": "R1", "name": "",
                           "from_room": "", "from_block": "A", "from_seat": "",
                           "to_room": "", "to_block": "B", "to_seat": "",
                           "from_cohort": "", "to_cohort": ""}])
    assert len(mailer.unmailed(trans, path=log)) == 1
    mailer.log_sends([dict(timestamp="2026-05-01T09:00:00", kind="seat_transition", roll="R1",
                           to="x", subject="s", status="sent", detail="")], log)
    assert len(mailer.unmailed(trans, path=log)) == 1, "a mail sent BEFORE the move counted"
    mailer.log_sends([dict(timestamp="2026-05-02T11:00:00", kind="seat_transition", roll="R1",
                           to="x", subject="s", status="handed-off", detail="")], log)
    assert mailer.unmailed(trans, path=log).empty


def test_gmail_compose_url_carries_everything():
    from urllib.parse import unquote
    # the block mail still signs with the full block, Devanagari included
    m = mailer.build_exam_block_mail(["B26CS1920"], exam="Quiz 1", when="", room="LHC 110",
                                     block="A")
    plain = unquote(mailer.compose_links([m], account="p24ph0999@example.edu",
                                         group=False)[0][1])
    assert "b26cs1920@example.edu" in plain
    assert "आपका नाम" in plain, "the Devanagari signature did not survive encoding"

    # the attendance notice signs off shorter, and carries the instructors in Cc
    att = mailer.build_attendance_mail("B26CS1920", name="A", present=8, absent=6, held=14,
                                       remaining=18)
    plain_att = unquote(mailer.compose_links([att], group=False)[0][1])
    assert mailer.DEFAULT_ATTENDANCE_CC in plain_att
    assert f"Teaching Assistant, {C.COURSE}" in plain_att


def test_identical_messages_batch_but_personal_ones_do_not():
    same = [mailer.build_transition_mail(f"B26XX{i:04d}", "C", name="student", room="LHC110")
            for i in range(5)]
    assert len(mailer.compose_links(same, to="office@example.edu")) == 1
    assert len(mailer.compose_links(same, group=False)) == 5


def test_app_password_despacing_leaves_real_passphrases_alone():
    assert mailer._clean_password("abcd efgh ijkl mnop") == "abcdefghijklmnop"
    assert mailer._clean_password("my long pass phrase") == "my long pass phrase"


def test_lhc105_exam_mail_shows_seat_number_not_row_col():
    url = "https://example.edu/find-my-seat.html"
    m = mailer.build_exam_mail("B25CI1930", room="LHC 105", block="R01", seat="15",
                               exam="Minor Exam", lookup_url=url)
    assert "Room LHC 105, Seat 15" in m.body
    assert "Block R01" not in m.body
    assert "row 1 col 15" not in m.body
    assert f"You can also check your position by entering roll number on this web page: {url}" in m.body


