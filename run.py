#!/usr/bin/env python3
"""LHC-110 attendance toolkit.

  plan    roster.xlsx  -> seat allocation, seating workbook, interactive hall plan
  sheets  roster.xlsx  -> printable signature-sheet PDFs + template.json

Typical term:

  python run.py plan   --roster roster_english.xlsx --out out/english
  python run.py sheets --roster roster_english.xlsx --out out/english/sheets

When the Hindi roster arrives, the same commands run against it:

  python run.py plan   --roster roster_hindi.xlsx --cohort "Hindi Batch" --out out/hindi

Filter one medium out of a combined roster with --medium Hindi.
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
import pandas as pd
import config as C
from seating import load_roster, allocate, allocate_grouped, block_summary
import reports


def _roster(a) -> pd.DataFrame:
    df = load_roster(a.roster, sheet=a.sheet, header=a.header)
    if a.medium:
        if "Medium" not in df:
            sys.exit("--medium given but the roster has no language/medium column.")
        df = df[df.Medium.str.lower() == a.medium.lower()].reset_index(drop=True)
        if df.empty:
            sys.exit(f"No students with medium {a.medium!r}.")
    return df


def _allocate(df: pd.DataFrame, a):
    """Plain roll-order allocate, or branch-grouped when --group is passed
    (see config.BRANCH_GROUPS - currently defined for LHC110 and LHC2 101,
    see config.py)."""
    if not getattr(a, "group", False):
        return allocate(df, a.room)
    groups = C.BRANCH_GROUPS.get(a.room)
    if not groups:
        sys.exit(f"--group given but no branch grouping is defined for room {a.room!r} "
                 f"in config.BRANCH_GROUPS.")
    return allocate_grouped(df, a.room, groups)


def cmd_plan(a):
    df = _roster(a)
    alloc, seats = _allocate(df, a)
    out = Path(a.out); out.mkdir(parents=True, exist_ok=True)
    alloc.to_csv(out / "allocation.csv", index=False)
    xlsx = reports.build_workbook(alloc, seats, str(out / f"{a.room}_seating.xlsx"),
                                  room=a.room, course=a.course, cohort=a.cohort,
                                  session=a.session, source=Path(a.roster).name)
    html = reports.build_html(alloc, seats, str(out / f"{a.room}_seating_plan.html"),
                              room=a.room, course=a.course, cohort=a.cohort, session=a.session)
    print(block_summary(alloc, seats, a.room).to_string(index=False))
    print(f"\n{len(alloc)} students seated")
    for p in (out / "allocation.csv", xlsx, html):
        print("  ->", p)


def cmd_sheets(a):
    import sheets as sheetmod
    df = _roster(a)
    alloc, _ = _allocate(df, a)
    tpl = sheetmod.build_all(alloc, a.out, course=a.course, room=a.room,
                             session=a.session, date_label=a.date or "", ta_name=a.ta or "")
    print(f"{len(alloc)} boxes across {len(alloc.Block.unique())} blocks")
    for p in sorted(Path(a.out).iterdir()):
        print("  ->", p)
    print("\nPrint at 100% scale (no 'fit to page').")


def main():
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)

    def common(q, roster=True):
        if roster:
            q.add_argument("--roster", required=True, help="Section & Group style .xlsx")
            q.add_argument("--sheet", default="Section & Group")
            q.add_argument("--header", type=int, default=1,
                           help="0-based row index of the header row (default 1)")
            q.add_argument("--medium", default=None,
                           help="keep only this medium, e.g. Hindi")
        q.add_argument("--room", default="LHC110")
        q.add_argument("--course", default=C.COURSE)
        q.add_argument("--cohort", default="English Batch")
        q.add_argument("--session", default=C.SESSION)
        q.add_argument("--group", action="store_true",
                       help="seat branch groups together per config.BRANCH_GROUPS "
                            "instead of plain roll order")

    q = sub.add_parser("plan", help="allocate seats and build the workbook and hall plan")
    common(q); q.add_argument("--out", default="out/plan"); q.set_defaults(f=cmd_plan)

    q = sub.add_parser("sheets", help="build printable signature sheets")
    common(q); q.add_argument("--out", default="out/sheets")
    q.add_argument("--date", default=None, help="printed on the sheet, e.g. '07 Aug 2026'")
    q.add_argument("--ta", default=None, help="attendance TA name(s) printed on the sheet")
    q.set_defaults(f=cmd_sheets)

    a = p.parse_args()
    a.f(a)


if __name__ == "__main__":
    main()
