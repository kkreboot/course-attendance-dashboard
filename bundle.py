#!/usr/bin/env python3
"""Every printable for one cohort or one exam, as a single ZIP.

The files are spread across `out/<cohort>/`, `out/<cohort>/sheets/`,
`out/exams/<slug>/` and a few in the project root (hall posters, the seat
allotment PDF, the find-your-block page). Printing for a class or an exam
used to mean collecting them one download button at a time; `cohort_bundle`
and `exam_bundle` collect them instead, with a MANIFEST.txt saying what each
file is and when it was built, so a stale sheet in the pile is visible before
it is printed.

Read-only over `out/`: nothing is regenerated here. Build first, then bundle.
"""
from __future__ import annotations

import io
import zipfile
from datetime import datetime
from pathlib import Path

import config as C

PRINTABLE = {".pdf", ".xlsx", ".html"}
OUT = Path("out")

# what a file is, by name, for the manifest
_KINDS = [
    ("ALL_signature_sheets", "all signature sheets, merged"),
    ("signature_block_", "signature sheet, one block"),
    ("_seating_plan.html", "interactive hall plan"),
    ("_seating.xlsx", "seating workbook"),
    ("poster_", "hall-door poster"),
    ("find-your-block", "find-your-block page (safe to publish)"),
    ("find-my-seat", "find-my-seat page (safe to publish)"),
    ("seat_map", "exam seat map"),
    ("_exam_pack", "invigilator pack"),
    ("_where_to_report", "where-to-report summary"),
    ("_notice_board", "notice-board map"),
    ("answer_showing", "answer-script showing"),
    ("Seating", "seat allotment PDF"),
]


def _kind(name: str) -> str:
    return next((k for pat, k in _KINDS if pat in name), "")


def _printables(folder: Path) -> list[Path]:
    if not folder.exists():
        return []
    return sorted(p for p in folder.rglob("*")
                  if p.is_file() and p.suffix.lower() in PRINTABLE
                  and not any(part.startswith(".") for part in p.relative_to(folder).parts))


def cohort_files(cohort_dir: str | Path, root: str | Path = ".") -> list[tuple[Path, str]]:
    """(file, name in the zip) for a classroom cohort: its sheets, workbook and
    hall plan, plus the shared root-level printables for its room."""
    cdir, root = Path(cohort_dir), Path(root)
    files = [(p, f"{cdir.name}/{p.relative_to(cdir).as_posix()}") for p in _printables(cdir)]
    rooms = {q.stem[: -len("_seating")] for q in cdir.glob("*_seating.xlsx")}
    shared = [root / f"poster_{r.replace(' ', '_')}.pdf" for r in sorted(rooms)]
    shared += sorted(root.glob(f"{C.COURSE_CODE}*Seating*.pdf"))
    shared += [root / "find-your-block.html"]
    files += [(p, f"shared/{p.name}") for p in shared if p.exists()]
    return files


def exam_files(exam_dir: str | Path) -> list[tuple[Path, str]]:
    """(file, name in the zip) for a saved exam: seat maps, lookup page, packs,
    posters and the answer-showing sheets, whatever has been built."""
    edir = Path(exam_dir)
    return [(p, f"{edir.name}/{p.relative_to(edir).as_posix()}") for p in _printables(edir)]


def stale(files: list[tuple[Path, str]], allocation: str | Path | None) -> list[str]:
    """Zip names of files built before `allocation` was last written - a
    sheet or seat map printed from them would show last week's seats."""
    if not allocation or not Path(allocation).exists():
        return []
    ref = Path(allocation).stat().st_mtime
    return [arc for p, arc in files if p.stat().st_mtime < ref - 1]


def manifest(files: list[tuple[Path, str]], title: str,
             allocation: str | Path | None = None) -> str:
    old = set(stale(files, allocation))
    lines = [f"{C.COURSE} - {title}",
             f"Bundled {datetime.now():%d %b %Y %H:%M}. Built dates are the files' own.",
             ("Files marked OLDER THAN ALLOCATION were built before the seats last changed: "
              "rebuild them before printing." if old else
              "Everything here was built after the allocation was last written."
              if allocation and Path(allocation).exists() else ""),
             ""]
    for p, arc in files:
        built = datetime.fromtimestamp(p.stat().st_mtime)
        kind = _kind(p.name)
        lines.append(f"{built:%d %b %H:%M}  {p.stat().st_size / 1024:8.0f} KB  {arc}"
                     + (f"  ({kind})" if kind else "")
                     + ("  ** OLDER THAN ALLOCATION **" if arc in old else ""))
    return "\n".join(lines) + "\n"


def make_zip(files: list[tuple[Path, str]], title: str,
             allocation: str | Path | None = None) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("MANIFEST.txt", manifest(files, title, allocation))
        for p, arc in files:
            z.write(p, arc)
    return buf.getvalue()


def cohort_bundle(cohort_dir: str | Path, root: str | Path = ".") -> tuple[bytes, int]:
    files = cohort_files(cohort_dir, root)
    return (make_zip(files, f"{Path(cohort_dir).name.title()} batch printables",
                     Path(cohort_dir) / "allocation.csv"), len(files))


def exam_bundle(exam_dir: str | Path, title: str = "") -> tuple[bytes, int]:
    files = exam_files(exam_dir)
    return (make_zip(files, f"{title or Path(exam_dir).name} printables",
                     Path(exam_dir) / "allocation.csv"), len(files))
