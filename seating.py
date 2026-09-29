"""Roster loading and seat allocation."""
from __future__ import annotations
import csv
import math
from datetime import datetime
from pathlib import Path

import pandas as pd
from config import ROOMS, RESERVE_WEDGE
import config as C

CANON = {
    "student rollno": "Roll", "roll no": "Roll", "rollno": "Roll", "roll": "Roll",
    "roll number": "Roll",
    "student name": "Name", "name": "Name",
    "program": "Program", "programme": "Program", "branch": "Program",
    "group": "Group", "language": "Medium", "medium": "Medium", "sr": "Sr",
}


_ROLL_KEYS = {"student rollno", "roll no", "rollno", "roll", "roll number"}


def _find_header_row(path: str, sheet: str, sample_rows: int = 15) -> int | None:
    """Scan the first few rows for one that has a roll-number-looking
    column, for rosters whose header isn't on the expected row (e.g. the
    Institute's own 'Subject Roll List' export puts it on row 3, under two
    title rows, not row 2 like the Section & Group workbooks)."""
    raw = pd.read_excel(path, sheet_name=sheet, header=None, nrows=sample_rows)
    for i, row in raw.iterrows():
        cells = {str(v).strip().lower() for v in row if pd.notna(v)}
        if cells & _ROLL_KEYS:
            return i
    return None


def load_roster(path: str, sheet: str = "Section & Group", header: int = 1) -> pd.DataFrame:
    """Read a roster workbook into Roll/Name/Program/Group/Medium.

    Tries the given sheet/header first (the Section & Group default), and
    falls back to auto-detecting both when they don't match - so a
    differently-shaped roster (e.g. the Institute's 'Subject Roll List'
    export, single sheet, header on row 3) can be pointed at directly
    without hand-tuning the sheet name and header row first."""
    try:
        book = pd.ExcelFile(path)
    except Exception:
        book = None
    sheet_names = book.sheet_names if book is not None else [sheet]
    use_sheet = sheet if sheet in sheet_names else (sheet_names[0] if len(sheet_names) == 1 else sheet)

    raw = pd.read_excel(path, sheet_name=use_sheet, header=header)
    has_roll = any(str(c).strip().lower() in _ROLL_KEYS for c in raw.columns)
    if not has_roll:
        detected = _find_header_row(path, use_sheet)
        if detected is not None and detected != header:
            raw = pd.read_excel(path, sheet_name=use_sheet, header=detected)

    ren = {c: CANON[str(c).strip().lower()] for c in raw.columns
           if str(c).strip().lower() in CANON}
    df = raw.rename(columns=ren)
    if "Roll" not in df:
        raise ValueError(f"No roll-number column found in {path!r}. Columns: {list(raw.columns)}")
    keep = [c for c in ["Roll", "Name", "Program", "Group", "Medium"] if c in df]
    df = df[keep].dropna(subset=["Roll"]).copy()
    df["Roll"] = df["Roll"].astype(str).str.strip()
    for c in keep:
        if c != "Roll":
            df[c] = df[c].astype(str).str.strip().replace({"nan": ""})
    dup = df.Roll[df.Roll.duplicated()].tolist()
    if dup:
        raise ValueError(f"Duplicate roll numbers in roster: {dup}")
    return df.sort_values("Roll").reset_index(drop=True)


ATTENDANCE_WORKBOOK = C.ATTENDANCE_WORKBOOK


def _cohort_sheets(path: str) -> list[str]:
    """Sheets that are one batch each. A sheet carrying a `Batch` column is the
    whole course on one list, and adding it to the two medium sheets counts
    every student twice."""
    out = []
    for name in pd.ExcelFile(path).sheet_names:
        header = _find_header_row(path, name)
        if header is None:
            continue
        cols = pd.read_excel(path, sheet_name=name, header=header, nrows=0).columns
        if any(str(c).strip().lower() == "batch" for c in cols):
            continue
        out.append(name)
    return out


def load_attendance_roster(path: str = ATTENDANCE_WORKBOOK,
                           sheet: str | None = None) -> pd.DataFrame:
    """Read the Institute's master Class Attendance workbook as a roster.

    This is the authoritative list of who is in which batch - one sheet per
    medium, kept current by the department - so it beats any Section & Group
    export, which goes stale the moment a student switches batch.

    Unlike `load_roster` the rows are returned in the workbook's **own order**
    (an `Order` column records it, per medium). That order is what the
    signature sheets follow, so a TA reading down a sheet reads down the same
    names in the same sequence as the attendance register.

    Pass `sheet` for one medium only; otherwise every *cohort* sheet is
    returned with a `Medium` column naming it.

    The workbook also carries a combined sheet: both batches on one list with a
    `Batch` column. That is a view of the same students, not a third cohort, so
    reading every sheet would return each student twice. It is skipped unless
    asked for by name, and asked for by name it takes its `Medium` from the
    `Batch` column rather than from the sheet name.
    """
    sheets = [sheet] if sheet else _cohort_sheets(path)
    out = []
    for name in sheets:
        header = _find_header_row(path, name)
        if header is None:
            raise ValueError(f"No roll-number column found on sheet {name!r} of {path!r}.")
        raw = pd.read_excel(path, sheet_name=name, header=header)
        ren = {c: CANON[str(c).strip().lower()] for c in raw.columns
               if str(c).strip().lower() in CANON}
        df = raw.rename(columns=ren)
        batch = next((c for c in df.columns if str(c).strip().lower() == "batch"), None)
        keep = [c for c in ("Roll", "Name") if c in df] + ([batch] if batch else [])
        df = df[keep].dropna(subset=["Roll"]).copy()
        df["Roll"] = df["Roll"].astype(str).str.strip()
        df = df[df.Roll != ""]
        if "Name" in df:
            df["Name"] = df["Name"].astype(str).str.strip()
        # the combined sheet says which batch each student is in; a medium
        # sheet says it by being that sheet
        df["Medium"] = (df[batch].astype(str).str.strip() if batch else name)
        if batch:
            df = df.drop(columns=[batch])
        df["Order"] = range(len(df))
        out.append(df.reset_index(drop=True))
    both = pd.concat(out, ignore_index=True)
    dup = both.Roll[both.Roll.duplicated()].tolist()
    if dup:
        raise ValueError(f"Duplicate roll numbers in {path!r}: {dup}")
    return both


def reorder_within_blocks(alloc: pd.DataFrame, roll_order) -> pd.DataFrame:
    """Re-sort each block's students into `roll_order`, seats unchanged.

    Nobody changes block and no seat changes occupancy - the same seats stay
    filled and the same seats stay empty. Only *which* student sits in each of
    a block's occupied seats changes, so that reading a block's seats in order
    reads its students in `roll_order` (the attendance register's order).

    A student missing from `roll_order` is placed within their branch in roll
    order if their branch exists in `roll_order`, otherwise keeps their
    relative position at the end of their block rather than being dropped.
    """
    roll_list = [str(r) for r in roll_order]
    rank = {r: i for i, r in enumerate(roll_list)}
    far = len(rank)

    def _rank_of(r: str, fallback_idx: int) -> float:
        if r in rank:
            return float(rank[r])
        b = branch_code(r)
        branch_rolls = [ro for ro in roll_list if branch_code(ro) == b]
        if branch_rolls:
            for prev in reversed(branch_rolls):
                if prev < r:
                    return rank[prev] + 0.5
            return rank[branch_rolls[0]] - 0.5
        return float(far + fallback_idx)

    out = []
    for block, sub in alloc.groupby("Block", sort=False):
        seats = sub.sort_values(["SeatRow", "SeatCol"])[["Seat", "SeatRow", "SeatCol"]]
        people = (sub.assign(_k=[_rank_of(str(r), i) for i, r in enumerate(sub.Roll)])
                     .sort_values("_k").drop(columns="_k").reset_index(drop=True))
        people[["Seat", "SeatRow", "SeatCol"]] = seats.reset_index(drop=True)
        out.append(people)
    return (pd.concat(out, ignore_index=True)
              .sort_values(["Block", "SeatRow", "SeatCol"]).reset_index(drop=True))


def seat_map(room: str = "LHC110") -> pd.DataFrame:
    """Every physical seat in the hall, with a stable code like 'E-014'."""
    rows = []
    for blk in ROOMS[room]:
        n = 0
        for r in range(1, blk.rows + 1):
            for c in range(1, blk.cols + 1):
                n += 1
                rows.append(dict(Block=blk.key, Section=blk.section, Row=r, Col=c,
                                 Kind="core", Seat=f"{blk.key}-{n:03d}"))
        for i, w in enumerate(blk.wedge, start=1):
            for c in range(1, w + 1):
                n += 1
                rows.append(dict(Block=blk.key, Section=blk.section, Row=blk.rows + i, Col=c,
                                 Kind="extra", Seat=f"{blk.key}-{n:03d}"))
    return pd.DataFrame(rows)


def _rear_row_plan(remaining: int, rear) -> dict:
    """Fill rear blocks in order, each to capacity in whole rows, before
    spilling into the next - same 'fewest blocks occupied' rule the front
    section uses, so a small overflow lands wholly in one rear block instead
    of being spread thin across all of them."""
    plan = {}
    for b in rear:
        rows = min(math.ceil(remaining / b.cols), b.rows) if remaining > 0 else 0
        plan[b.key] = rows
        remaining -= rows * b.cols
    if remaining > 0:
        raise ValueError(f"Hall too small: {remaining} students do not fit in the rear blocks.")
    return plan


def allocate(roster: pd.DataFrame, room: str = "LHC110") -> tuple[pd.DataFrame, pd.DataFrame]:
    """Assign seats in roll order. Front core seats fill first, wedge seats stay
    vacant as the transition pool, then rear blocks fill by whole rows.

    Returns (allocation, seats) where `seats` carries an `Occupied` flag.
    """
    blocks = ROOMS[room]
    front = [b for b in blocks if b.section == "front"]
    rear = [b for b in blocks if b.section == "rear"]
    seats = seat_map(room)

    n = len(roster)
    front_cap = sum(b.core for b in front)
    plan = {b.key: b.rows for b in front}
    plan.update(_rear_row_plan(n - front_cap, rear) if n > front_cap else {b.key: 0 for b in rear})
    if n < front_cap:                      # small cohort: trim front from the back
        short = front_cap - n
        for b in reversed(front):
            rows_drop = min(short // b.cols, plan[b.key])
            plan[b.key] -= rows_drop
            short -= rows_drop * b.cols

    usable = (seats.Kind.eq("core") if RESERVE_WEDGE else True) & \
             seats.apply(lambda r: r.Row <= plan[r.Block], axis=1)
    order = {b.key: i for i, b in enumerate(blocks)}
    avail = (seats[usable]
             .assign(_o=lambda d: d.Block.map(order))
             .sort_values(["_o", "Row", "Col"])
             .drop(columns="_o").reset_index(drop=True))
    if len(avail) < n:
        raise ValueError(f"Only {len(avail)} seats available for {n} students in {room}.")

    alloc = roster.copy().reset_index(drop=True)
    take = avail.iloc[:n].reset_index(drop=True)
    alloc["Block"] = take.Block
    alloc["Seat"] = take.Seat
    alloc["SeatRow"] = take.Row.astype(int)
    alloc["SeatCol"] = take.Col.astype(int)
    seats["Occupied"] = seats.Seat.isin(set(alloc.Seat))
    return alloc, seats


def branch_code(roll: str) -> str:
    """2-letter department code from a current-year roll number, e.g.
    'B26AE1901' -> 'AE'. Backlog students (any other admission year) aren't
    broken into branches - the arrangement sheet lumps them as one total per
    admission year, so they're grouped under 'BACKLOG' instead."""
    roll = str(roll)
    if roll.startswith("B26") and len(roll) >= 5:
        return roll[3:5]
    return "BACKLOG"


def allocate_grouped(roster: pd.DataFrame, room: str,
                     block_groups: dict[str, list[str]]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Like `allocate`, but seats branch groups together where specified:
    `block_groups` maps a block key to the branch codes reserved for it
    (see `config.BRANCH_GROUPS`). A block with no entry, and any student
    whose branch isn't listed for any block (including backlog students,
    who are never split by branch), falls through to a catch-all pass that
    fills remaining capacity in the room's normal front-to-rear order -
    the same fallback `allocate` always uses.

    Only front-of-room core capacity is used for the grouped pass, matching
    `allocate`'s convention of holding wedge seats in reserve; rear blocks
    participate in grouping (or catch-all) at their full row x col capacity,
    since rear blocks have no wedge.
    """
    blocks = ROOMS[room]
    seats = seat_map(room)

    roster = roster.copy().reset_index(drop=True)
    roster["_branch"] = roster.Roll.map(branch_code)

    cap_left = {b.key: b.core for b in blocks}
    block_of = pd.Series(pd.NA, index=roster.index, dtype=object)

    for b in blocks:
        wanted = block_groups.get(b.key)
        if not wanted:
            continue
        cand = roster[roster._branch.isin(wanted) & block_of.isna()].sort_values("Roll")
        take_idx = cand.index[: cap_left[b.key]]
        block_of.loc[take_idx] = b.key
        cap_left[b.key] -= len(take_idx)

    remaining = roster[block_of.isna()].sort_values("Roll")
    if len(remaining) > sum(cap_left.values()):
        raise ValueError(f"Only {sum(cap_left.values())} seats left for {len(remaining)} "
                         f"students in {room} once branch groups are placed.")
    for idx in remaining.index:
        b = next(b for b in blocks if cap_left[b.key] > 0)
        block_of.at[idx] = b.key
        cap_left[b.key] -= 1

    roster["Block"] = block_of
    alloc_rows = []
    for b in blocks:
        sub = roster[roster.Block == b.key]
        blk_seats = (seats[(seats.Block == b.key) & (seats.Kind == "core")]
                     .sort_values(["Row", "Col"]))
        for (_, srow), (_, seat) in zip(sub.iterrows(), blk_seats.iterrows()):
            row = srow.drop("_branch").to_dict()
            row.update(Seat=seat.Seat, SeatRow=int(seat.Row), SeatCol=int(seat.Col))
            alloc_rows.append(row)

    alloc = pd.DataFrame(alloc_rows).sort_values("Roll").reset_index(drop=True)
    seats["Occupied"] = seats.Seat.isin(set(alloc.Seat))
    return alloc, seats


def vacant_seats(alloc: pd.DataFrame, room: str, block: str, include_wedge: bool = False) -> pd.DataFrame:
    """Seats in one block of `room` not already taken in `alloc`, row order.
    Core seats only by default - same reserve-pool convention `allocate`
    uses, so a manual move can't quietly eat into the transition pool a late
    admission might need. Pass `include_wedge=True` to draw from that pool
    on purpose (e.g. a bulk batch transfer that needs the extra capacity)."""
    seats = seat_map(room)
    kinds = {"core", "extra"} if include_wedge else {"core"}
    occupied = set(alloc.Seat)
    return (seats[(seats.Block == block) & (seats.Kind.isin(kinds)) & (~seats.Seat.isin(occupied))]
            .sort_values(["Row", "Col"]).reset_index(drop=True))


def compact_block(alloc: pd.DataFrame, room: str, block: str) -> pd.DataFrame:
    """Close any seat number gaps in `block` of `room`, keeping relative student order.

    Vacant capacity belongs at the end of each block, contiguous from 1.
    """
    sub = alloc[alloc.Block == block].sort_values(["SeatRow", "SeatCol"]).copy().reset_index(drop=True)
    if sub.empty:
        return alloc
    seats = seat_map(room)
    kinds = {"core", "extra"}
    blk_seats = (seats[(seats.Block == block) & (seats.Kind.isin(kinds))]
                 .sort_values(["Row", "Col"]).reset_index(drop=True))
    for i in range(len(sub)):
        s = blk_seats.iloc[i]
        sub.at[i, "Seat"] = s.Seat
        sub.at[i, "SeatRow"] = int(s.Row)
        sub.at[i, "SeatCol"] = int(s.Col)
    other = alloc[alloc.Block != block]
    return (pd.concat([other, sub], ignore_index=True)
            .sort_values(["Block", "SeatRow", "SeatCol"]).reset_index(drop=True))


def reorder_block_in_roll_order(alloc: pd.DataFrame, room: str, block: str,
                                mode: str = "register",
                                reg_order: list[str] | None = None) -> pd.DataFrame:
    """Re-sort students in `block` of `room` into roll order without changing blocks.

    Seats run contiguous from 1 to N, with vacant capacity at the end.
    Modes:
      - 'register' (default): Follows the attendance register's row order from
        the master Class Attendance workbook. Any newly moved student whose roll
        isn't in the register yet is slotted right into their branch in roll order.
      - 'branch_roll': Preserves the block's branch grouping (e.g. CM,
        CY, EE in LHC110 Block C). Students within each branch are sorted by roll
        number.
      - 'roll': Pure alphabetical/numerical sort by roll number.
    """
    sub = alloc[alloc.Block == block].copy()
    if sub.empty:
        return alloc

    seats = seat_map(room)
    kinds = {"core", "extra"}
    blk_seats = (seats[(seats.Block == block) & (seats.Kind.isin(kinds))]
                 .sort_values(["Row", "Col"]).reset_index(drop=True))

    if mode == "roll":
        sub_sorted = sub.sort_values("Roll").reset_index(drop=True)
    elif mode == "register":
        if reg_order is None:
            # Every batch sheet, in workbook order. Roll numbers are unique
            # across batches, so ranking over the union orders this block
            # exactly as its own sheet would - without assuming which room
            # belongs to which sheet (this used to be `"English" if room ==
            # "LHC110" else "Hindi"`, i.e. only ever right for the demo).
            try:
                r_df = load_attendance_roster(C.ATTENDANCE_WORKBOOK)
                reg_order = list(r_df.Roll)
            except Exception:
                reg_order = None

        if reg_order is not None:
            roll_list = [str(r) for r in reg_order]
            rank = {r: i for i, r in enumerate(roll_list)}
            far = len(rank)

            def _rank_of(r: str, fallback_idx: int) -> float:
                if r in rank:
                    return float(rank[r])
                b = branch_code(r)
                branch_rolls = [ro for ro in roll_list if branch_code(ro) == b]
                if branch_rolls:
                    for prev in reversed(branch_rolls):
                        if prev < r:
                            return rank[prev] + 0.5
                    return rank[branch_rolls[0]] - 0.5
                return float(far + fallback_idx)

            sub["_k"] = [_rank_of(str(r), i) for i, r in enumerate(sub.Roll)]
            sub_sorted = (sub.sort_values("_k").drop(columns=["_k"]).reset_index(drop=True))
        else:
            mode = "branch_roll"

    if mode == "branch_roll":
        # 'branch_roll': discover the branch sequence of the block from seated order
        present_branches = []
        for r in sub.sort_values(["SeatRow", "SeatCol"]).Roll:
            b = branch_code(r)
            if b not in present_branches:
                present_branches.append(b)

        b_map = {b: i for i, b in enumerate(present_branches)}
        sub["_b_rank"] = sub.Roll.map(branch_code).map(lambda b: b_map.get(b, 999))
        sub_sorted = (sub.sort_values(["_b_rank", "Roll"])
                      .drop(columns=["_b_rank"])
                      .reset_index(drop=True))

    for i in range(len(sub_sorted)):
        s = blk_seats.iloc[i]
        sub_sorted.at[i, "Seat"] = s.Seat
        sub_sorted.at[i, "SeatRow"] = int(s.Row)
        sub_sorted.at[i, "SeatCol"] = int(s.Col)

    other = alloc[alloc.Block != block]
    return (pd.concat([other, sub_sorted], ignore_index=True)
            .sort_values(["Block", "SeatRow", "SeatCol"]).reset_index(drop=True))


def reorder_allocation(alloc: pd.DataFrame, room: str, blocks: list[str] | None = None,
                       mode: str = "register") -> pd.DataFrame:
    """Reorder one or more blocks in `alloc` so seats follow roll order."""
    if blocks is None:
        blocks = [b.key for b in ROOMS[room]]
    df = alloc.copy()
    for b in blocks:
        df = reorder_block_in_roll_order(df, room, b, mode=mode)
    return df


def move_student(alloc_src: pd.DataFrame, room_src: str, roll: str, room_dst: str, block_dst: str,
                 alloc_dst: pd.DataFrame | None = None, seat_dst: str | None = None,
                 include_wedge: bool = False, reorder: bool = False,
                 order_mode: str = "register") -> tuple[pd.DataFrame, pd.DataFrame, str]:
    """Move one student into a vacant seat of `block_dst` in `room_dst` -
    the same room for a block-to-block move (pass `alloc_dst=None`), or a
    different one for a cross-batch transfer (pass that batch's own
    allocation). `seat_dst` pins a specific seat; omitted, the first vacant
    seat in room order is used. Core seats only unless `include_wedge=True`
    (see `vacant_seats`) - wedge/extra seats are the transition pool and
    stay untouched by default. If `reorder=True`, the destination block is
    re-sorted into roll order (keeping blocks unchanged) and the source block
    is compacted to close any gaps.

    Returns (new_alloc_src, new_alloc_dst, seat_used) - for a same-room move
    these are the same (single, already-merged) DataFrame. Any attendance
    already recorded under the student's old seat is *not* retroactively
    moved - this only changes where they sit from here on.
    """
    if roll not in set(alloc_src.Roll):
        raise ValueError(f"{roll!r} not found in the source allocation.")
    if room_dst not in ROOMS:
        raise ValueError(f"Unknown room {room_dst!r}.")
    if block_dst not in {b.key for b in ROOMS[room_dst]}:
        raise ValueError(f"Block {block_dst!r} does not exist in {room_dst!r}.")

    row = alloc_src.loc[alloc_src.Roll == roll].iloc[0].copy()
    src_block = str(row["Block"])
    remaining_src = alloc_src[alloc_src.Roll != roll].reset_index(drop=True)
    same_room_move = alloc_dst is None
    base_dst = remaining_src if same_room_move else alloc_dst

    if seat_dst:
        seats = seat_map(room_dst)
        kinds = {"core", "extra"} if include_wedge else {"core"}
        match = seats[(seats.Block == block_dst) & (seats.Kind.isin(kinds)) & (seats.Seat == seat_dst)]
        if match.empty:
            kind_note = "core or transition" if include_wedge else "core"
            raise ValueError(f"{seat_dst!r} is not a {kind_note} seat in block {block_dst} of {room_dst}.")
        if seat_dst in set(base_dst.Seat):
            raise ValueError(f"Seat {seat_dst} is already occupied.")
        pick = match.iloc[0]
    else:
        avail = vacant_seats(base_dst, room_dst, block_dst, include_wedge=include_wedge)
        if avail.empty:
            kind_note = "core or transition" if include_wedge else "core"
            raise ValueError(f"No vacant {kind_note} seats left in block {block_dst} of {room_dst}.")
        pick = avail.iloc[0]

    row["Block"] = block_dst
    row["Seat"] = pick.Seat
    row["SeatRow"] = int(pick.Row)
    row["SeatCol"] = int(pick.Col)
    new_dst = pd.concat([base_dst, row.to_frame().T], ignore_index=True)
    new_dst["Roll"] = new_dst.Roll.astype(str)
    new_dst = new_dst.sort_values("Roll").reset_index(drop=True)

    if reorder and not seat_dst:
        new_dst = reorder_block_in_roll_order(new_dst, room_dst, block_dst, mode=order_mode)
        actual_seat = str(new_dst.loc[new_dst.Roll == roll, "Seat"].iloc[0])
        if same_room_move:
            if src_block != block_dst:
                new_dst = compact_block(new_dst, room_src, src_block)
            return new_dst, new_dst, actual_seat
        else:
            remaining_src = compact_block(remaining_src, room_src, src_block)
            return remaining_src, new_dst, actual_seat

    if same_room_move:
        return new_dst, new_dst, pick.Seat
    return remaining_src, new_dst, pick.Seat


def block_summary(alloc: pd.DataFrame, seats: pd.DataFrame, room: str = "LHC110") -> pd.DataFrame:
    out = []
    for b in ROOMS[room]:
        sub = alloc[alloc.Block == b.key]
        out.append(dict(Block=b.key, Position=b.side, PerRow=b.cols, Core=b.core,
                        Wedge=b.extra, Capacity=b.total, Seated=len(sub),
                        RowsUsed=int(sub.SeatRow.max()) if len(sub) else 0, RowsTotal=b.rows,
                        Vacant=b.total - len(sub),
                        FirstRoll=sub.Roll.min() if len(sub) else "",
                        LastRoll=sub.Roll.max() if len(sub) else ""))
    return pd.DataFrame(out)


# ─────────────────────────── transition history ──────────────────────────
# `allocation.csv` records where everyone sits *now*; it keeps no history, so
# after a move there is no way to answer "who changed block this week?" - and
# that is exactly the set of students who need telling. Every move through the
# dashboard appends a row here.
TRANSITION_LOG = Path("out") / "transitions.csv"
TRANSITION_COLS = ["timestamp", "roll", "name", "from_room", "from_block",
                   "from_seat", "to_room", "to_block", "to_seat",
                   "from_cohort", "to_cohort"]


def log_transition(roll: str, *, name: str = "", from_room: str = "",
                   from_block: str = "", from_seat: str = "", to_room: str = "",
                   to_block: str = "", to_seat: str = "", from_cohort: str = "",
                   to_cohort: str = "", path: str | Path = TRANSITION_LOG) -> Path:
    """Append one move to `out/transitions.csv`.

    Kept deliberately append-only and separate from `allocation.csv`: the
    allocation is state and gets rewritten wholesale by every regeneration,
    while this is history and must survive that.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    new = not path.exists()
    with path.open("a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=TRANSITION_COLS)
        if new:
            w.writeheader()
        w.writerow({
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "roll": str(roll), "name": name,
            "from_room": from_room, "from_block": from_block, "from_seat": from_seat,
            "to_room": to_room, "to_block": to_block, "to_seat": to_seat,
            "from_cohort": from_cohort, "to_cohort": to_cohort,
        })
    return path


def load_transitions(path: str | Path = TRANSITION_LOG) -> pd.DataFrame:
    """Every logged move, newest last. Empty (but correctly-shaped) frame if
    nothing has been moved yet, so callers can filter without a None check."""
    path = Path(path)
    if not path.exists():
        return pd.DataFrame(columns=TRANSITION_COLS)
    df = pd.read_csv(path, dtype=str).fillna("")
    for c in TRANSITION_COLS:
        if c not in df.columns:
            df[c] = ""
    return df[TRANSITION_COLS]


def block_changed(df: pd.DataFrame | None = None) -> pd.DataFrame:
    """Moves that actually changed *block* - the only ones worth mailing about.

    A move within one block (a seat renumber, a reorder, a compaction) changes
    `Seat` but not where the student sits at block granularity, and mail here
    never names a seat. Telling someone their block changed when it didn't is
    worse than saying nothing.
    """
    df = load_transitions() if df is None else df
    if df.empty:
        return df
    changed = (df.from_block != df.to_block) | (df.from_room != df.to_room)
    return df[changed].reset_index(drop=True)


def detect_untracked_moves(cohort: str, alloc_csv: str | Path | None = None) -> pd.DataFrame:
    """Block changes that happened *outside* this toolkit.

    `log_transition` only sees moves made through the dashboard. A block edited
    by hand in Excel, or a row changed straight in `allocation.csv`, leaves no
    row - so those students never appear in "who still needs telling", which is
    exactly the list that matters.

    This diffs the live `allocation.csv` against the copy kept in
    `out/.history/state/` from the last time the toolkit wrote or acknowledged
    it. Returns the same columns `load_transitions` uses, so the two can be
    concatenated. Empty frame when there's no baseline yet - a first run has
    nothing to compare against, and inventing moves from nothing would be
    worse than reporting none.
    """
    import history

    cols = TRANSITION_COLS
    base_path = history.state_path(cohort)
    alloc_csv = Path(alloc_csv or Path("out") / cohort / "allocation.csv")
    if not base_path.exists() or not alloc_csv.exists():
        return pd.DataFrame(columns=cols)

    base = pd.read_csv(base_path, dtype={"Roll": str}).set_index("Roll")
    now = pd.read_csv(alloc_csv, dtype={"Roll": str}).set_index("Roll")
    stamp = datetime.fromtimestamp(alloc_csv.stat().st_mtime).isoformat(timespec="seconds")

    rows = []
    for roll in now.index.intersection(base.index):
        old_block, new_block = str(base.at[roll, "Block"]), str(now.at[roll, "Block"])
        if old_block == new_block:
            continue
        rows.append({
            "timestamp": stamp, "roll": roll,
            "name": str(now.at[roll, "Name"]) if "Name" in now.columns else "",
            "from_room": "", "from_block": old_block,
            "from_seat": str(base.at[roll, "Seat"]) if "Seat" in base.columns else "",
            "to_room": "", "to_block": new_block,
            "to_seat": str(now.at[roll, "Seat"]) if "Seat" in now.columns else "",
            "from_cohort": cohort, "to_cohort": cohort,
        })
    # A student who only appears on one side moved *between* cohorts or joined
    # /left entirely; that isn't a block change within this file and is left to
    # the health check, which looks at both cohorts and the roster at once.
    return pd.DataFrame(rows, columns=cols)


def adopt_untracked_moves(cohort: str, alloc_csv: str | Path | None = None) -> int:
    """Write detected hand-made moves into `transitions.csv` and re-baseline,
    so they show up in the mail page like any other move. Returns the count."""
    import history

    found = detect_untracked_moves(cohort, alloc_csv)
    alloc_csv = Path(alloc_csv or Path("out") / cohort / "allocation.csv")
    for r in found.itertuples():
        log_transition(r.roll, name=r.name, from_room=r.from_room, from_block=r.from_block,
                       from_seat=r.from_seat, to_room=r.to_room, to_block=r.to_block,
                       to_seat=r.to_seat, from_cohort=r.from_cohort, to_cohort=r.to_cohort)
    if alloc_csv.exists():
        history.save_state(cohort, alloc_csv)
    return len(found)
