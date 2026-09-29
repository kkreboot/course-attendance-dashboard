"""Invariants the seating code must never break.

These were all verified by hand at least once during development, which is
exactly why they belong here instead: a rule checked by whoever remembers it
is a rule that eventually isn't.
"""
import pandas as pd
import pytest

import config as C
import seating


def alloc_for(room="LHC110", n=40, block="A"):
    seats = seating.seat_map(room)
    seats = seats[(seats.Block == block) & (seats.Kind == "core")].head(n)
    return pd.DataFrame({
        "Roll": [f"B26XX{i:04d}" for i in range(1, len(seats) + 1)],
        "Name": [f"Student {i}" for i in range(1, len(seats) + 1)],
        "Medium": "English", "Block": block,
        "Seat": list(seats.Seat), "SeatRow": list(seats.Row), "SeatCol": list(seats.Col),
    })


def test_seat_map_is_contiguous_and_ordered():
    for room in C.ROOMS:
        sm = seating.seat_map(room)
        for block, sub in sm.groupby("Block"):
            sub = sub.sort_values(["Row", "Col"])
            assert list(sub.Seat) == sorted(sub.Seat), f"{room} {block} seat ids out of row order"
        assert not sm.Seat.duplicated().any(), f"{room} has duplicate seat ids"


def test_reorder_keeps_block_and_occupancy():
    a = alloc_for(n=20)
    new_order = list(reversed(a.Roll))
    out = seating.reorder_within_blocks(a, new_order)
    assert set(out.Roll) == set(a.Roll)
    assert set(out.Seat) == set(a.Seat), "reorder changed which seats are occupied"
    assert out.set_index("Roll").Block.to_dict() == a.set_index("Roll").Block.to_dict()
    assert list(out.sort_values(["SeatRow", "SeatCol"]).Roll) == new_order


def test_reorder_keeps_unlisted_students():
    a = alloc_for(n=10)
    partial = list(a.Roll[:5])
    out = seating.reorder_within_blocks(a, partial)
    assert set(out.Roll) == set(a.Roll), "a student missing from the order was dropped"


def test_move_student_uses_core_seats_and_frees_the_old_one():
    a = alloc_for(n=10)
    before = a.loc[a.Roll == "B26XX0003"].iloc[0]
    out, _, seat = seating.move_student(a, "LHC110", "B26XX0003", "LHC110", "B")
    row = out.loc[out.Roll == "B26XX0003"].iloc[0]
    assert row.Block == "B" and row.Seat == seat
    assert before.Seat not in set(out.Seat), "the vacated seat is still marked occupied"
    assert len(out) == len(a)


def test_move_student_rejects_unknown_roll_and_block():
    a = alloc_for(n=5)
    with pytest.raises(ValueError):
        seating.move_student(a, "LHC110", "NOPE", "LHC110", "B")
    with pytest.raises(ValueError):
        seating.move_student(a, "LHC110", "B26XX0001", "LHC110", "Z")


def test_wedge_seats_are_held_back_unless_asked_for():
    """Front blocks with a taper keep their extra seats as the transition pool."""
    room, block = "LHC110", "B"
    seats = seating.seat_map(room)
    core = seats[(seats.Block == block) & (seats.Kind == "core")]
    full_b = alloc_for(room, n=len(core), block=block)      # every core seat in B taken
    assert seating.vacant_seats(full_b, room, block).empty
    assert not seating.vacant_seats(full_b, room, block, include_wedge=True).empty

    # someone sitting in A tries to move into the full block B
    mover = alloc_for(room, n=1, block="A")
    mover["Roll"] = ["B26YY0001"]
    both = pd.concat([full_b, mover], ignore_index=True)
    with pytest.raises(ValueError, match="No vacant"):
        seating.move_student(both, room, "B26YY0001", room, block)
    # ... and succeeds when the transition pool is explicitly opened up
    out, _, seat = seating.move_student(both, room, "B26YY0001", room, block,
                                        include_wedge=True)
    assert out.loc[out.Roll == "B26YY0001"].iloc[0].Block == block


def test_transition_log_and_block_changed(tmp_path):
    log = tmp_path / "t.csv"
    seating.log_transition("R1", from_block="A", to_block="B", path=log)
    seating.log_transition("R2", from_block="C", to_block="C", path=log)   # seat renumber only
    df = seating.load_transitions(log)
    assert len(df) == 2
    changed = seating.block_changed(df)
    assert list(changed.roll) == ["R1"], "a move inside one block was reported as a block change"


def test_compact_block_closes_gaps():
    a = alloc_for(room="LHC110", n=5, block="A")
    # Drop seat 2 (A-002) to create a gap
    a_gap = a[a.Roll != "B26XX0002"].copy()
    compacted = seating.compact_block(a_gap, "LHC110", "A")
    sub = compacted[compacted.Block == "A"].sort_values(["SeatRow", "SeatCol"])
    assert list(sub.Seat) == ["A-001", "A-002", "A-003", "A-004"]
    assert list(sub.Roll) == ["B26XX0001", "B26XX0003", "B26XX0004", "B26XX0005"]


def test_reorder_block_in_roll_order_branch_roll():
    # Mix branches: CM and CY. CY student placed after EE
    sm = seating.seat_map("LHC110")
    c_seats = sm[(sm.Block == "C") & (sm.Kind == "core")].head(5).reset_index(drop=True)
    df = pd.DataFrame({
        "Roll": ["B26CM1901", "B26CY1902", "B26EE1901", "B26EE1902", "B26CY1915"],
        "Name": ["S1", "S2", "S3", "S4", "S5"],
        "Block": "C", "Medium": "English",
        "Seat": list(c_seats.Seat), "SeatRow": list(c_seats.Row), "SeatCol": list(c_seats.Col),
    })
    reordered = seating.reorder_block_in_roll_order(df, "LHC110", "C", mode="branch_roll")
    sub = reordered.sort_values(["SeatRow", "SeatCol"])
    # B26CY1915 should be right after B26CY1902, before B26EE1901
    assert list(sub.Roll) == ["B26CM1901", "B26CY1902", "B26CY1915", "B26EE1901", "B26EE1902"]
    assert list(sub.Seat) == list(c_seats.Seat)


def test_reorder_block_in_roll_order_pure_roll():
    sm = seating.seat_map("LHC110")
    c_seats = sm[(sm.Block == "C") & (sm.Kind == "core")].head(3).reset_index(drop=True)
    df = pd.DataFrame({
        "Roll": ["B26EE1901", "B26CM1901", "B26CY1501"],
        "Name": ["S1", "S2", "S3"],
        "Block": "C", "Medium": "English",
        "Seat": list(c_seats.Seat), "SeatRow": list(c_seats.Row), "SeatCol": list(c_seats.Col),
    })
    reordered = seating.reorder_block_in_roll_order(df, "LHC110", "C", mode="roll")
    sub = reordered.sort_values(["SeatRow", "SeatCol"])
    assert list(sub.Roll) == ["B26CM1901", "B26CY1501", "B26EE1901"]


def test_move_student_with_reorder():
    sm = seating.seat_map("LHC110")
    a_seats = sm[(sm.Block == "A") & (sm.Kind == "core")].head(4).reset_index(drop=True)
    b_seats = sm[(sm.Block == "B") & (sm.Kind == "core")].head(3).reset_index(drop=True)
    alloc_a = pd.DataFrame({
        "Roll": ["B26BB1901", "B26BB1502", "B26BB1503", "B26BB1904"],
        "Name": ["A1", "A2", "A3", "A4"], "Block": "A", "Medium": "English",
        "Seat": list(a_seats.Seat), "SeatRow": list(a_seats.Row), "SeatCol": list(a_seats.Col),
    })
    alloc_b = pd.DataFrame({
        "Roll": ["B26ME1501", "B26ME1903"],
        "Name": ["B1", "B3"], "Block": "B", "Medium": "English",
        "Seat": list(b_seats.Seat[:2]), "SeatRow": list(b_seats.Row[:2]), "SeatCol": list(b_seats.Col[:2]),
    })
    both = pd.concat([alloc_a, alloc_b], ignore_index=True)
    # Move B26BB1502 to Block B with a middle roll number B26ME1902 (simulated by creating roll)
    # Give mover roll B26ME1902
    both.loc[both.Roll == "B26BB1502", "Roll"] = "B26ME1902"
    both.loc[both.Roll == "B26ME1902", "Name"] = "B2"

    new_src, new_dst, seat_used = seating.move_student(
        both, "LHC110", "B26ME1902", "LHC110", "B", reorder=True
    )
    # Source Block A should be compacted (3 students, contiguous from A-001)
    sub_a = new_dst[new_dst.Block == "A"].sort_values(["SeatRow", "SeatCol"])
    assert list(sub_a.Seat) == ["A-001", "A-002", "A-003"]
    # Destination Block B should be sorted in roll order (ME1001, ME1002, ME1003)
    sub_b = new_dst[new_dst.Block == "B"].sort_values(["SeatRow", "SeatCol"])
    assert list(sub_b.Roll) == ["B26ME1501", "B26ME1902", "B26ME1903"]
    assert list(sub_b.Seat) == ["B-001", "B-002", "B-003"]
    assert seat_used == "B-002"
