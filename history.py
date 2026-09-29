#!/usr/bin/env python3
"""Rolling snapshots of the files this toolkit overwrites, so a bad build can
be undone.

Why not `.bak` files next to the originals: that convention was tried earlier
in this project's life and explicitly abandoned - `*.pre-something.xlsx` piled
up in the working folder until nobody could tell which copy was current. This
keeps the same safety in one hidden, capped, git-ignored place instead:
`out/.history/<timestamp>_<label>.zip`, newest `KEEP` retained, everything
older pruned automatically.

The trigger for this was real: a build once wrote one cohort's students into
the other cohort's `sheets/` folder and deleted two blocks' PDFs, and the only
reason it was recoverable was that the source allocation still existed. A
snapshot taken immediately before each write makes that a one-click undo
rather than a rebuild-and-hope.
"""
from __future__ import annotations

import hashlib
import json
import shutil
import zipfile
from datetime import datetime
from pathlib import Path

HISTORY_DIR = Path("out") / ".history"
STATE_DIR = HISTORY_DIR / "state"
KEEP = 20
MANIFEST = "_manifest.json"


def _sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()[:16]


def _expand(paths) -> list[Path]:
    """Directories become their files; missing paths are skipped, since a
    snapshot taken before the *first* build of something has nothing to save."""
    out: list[Path] = []
    for p in paths:
        p = Path(p)
        if p.is_dir():
            out += [q for q in sorted(p.rglob("*")) if q.is_file()]
        elif p.is_file():
            out.append(p)
    return out


def snapshot(paths, label: str, note: str = "") -> Path | None:
    """Zip `paths` as they are right now. Returns None if there was nothing to
    save - a caller shouldn't have to check whether files exist yet."""
    files = _expand(paths)
    if not files:
        return None
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    safe = "".join(c if c.isalnum() or c in "-_" else "-" for c in label)[:40]
    zpath = HISTORY_DIR / f"{stamp}_{safe}.zip"
    manifest = {"created": datetime.now().isoformat(timespec="seconds"),
                "label": label, "note": note,
                "files": [{"path": str(f), "sha": _sha(f), "bytes": f.stat().st_size}
                          for f in files]}
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(f, str(f))
        z.writestr(MANIFEST, json.dumps(manifest, indent=1))
    prune()
    return zpath


def list_snapshots() -> list[dict]:
    """Newest first. A zip that can't be read is reported rather than skipped -
    a silently-ignored corrupt snapshot is worse than a visible broken one."""
    if not HISTORY_DIR.exists():
        return []
    out = []
    for z in sorted(HISTORY_DIR.glob("*.zip"), reverse=True):
        row = {"path": z, "name": z.name, "bytes": z.stat().st_size,
               "created": "", "label": "", "note": "", "files": [], "ok": True}
        try:
            with zipfile.ZipFile(z) as zf:
                m = json.loads(zf.read(MANIFEST))
            row.update(created=m.get("created", ""), label=m.get("label", ""),
                       note=m.get("note", ""), files=[f["path"] for f in m.get("files", [])])
        except (zipfile.BadZipFile, KeyError, json.JSONDecodeError, OSError) as e:
            row.update(ok=False, label=f"unreadable ({type(e).__name__})")
        out.append(row)
    return out


def restore(zip_path, only: list[str] | None = None, dry_run: bool = False) -> list[str]:
    """Put a snapshot's files back where they came from.

    Takes its own snapshot of whatever it is about to overwrite first - undoing
    an undo has to be possible too, or this just moves the danger one step.
    """
    zip_path = Path(zip_path)
    with zipfile.ZipFile(zip_path) as zf:
        names = [n for n in zf.namelist() if n != MANIFEST]
        if only:
            names = [n for n in names if n in set(only)]
        if dry_run:
            return names
        snapshot([n for n in names if Path(n).exists()], "before-restore",
                 note=f"replaced by {zip_path.name}")
        for n in names:
            dest = Path(n)
            dest.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(n) as src, dest.open("wb") as dst:
                shutil.copyfileobj(src, dst)
    return names


def prune(keep: int = KEEP) -> int:
    """Drop the oldest snapshots past `keep`. Returns how many were removed."""
    zips = sorted(HISTORY_DIR.glob("*.zip"), reverse=True) if HISTORY_DIR.exists() else []
    dropped = 0
    for z in zips[keep:]:
        z.unlink()
        dropped += 1
    return dropped


def total_bytes() -> int:
    return sum(z.stat().st_size for z in HISTORY_DIR.glob("*.zip")) if HISTORY_DIR.exists() else 0


# ─────────────────────── last-known state (drift detection) ───────────────────
def state_path(cohort: str) -> Path:
    return STATE_DIR / f"{cohort}_allocation.csv"


def save_state(cohort: str, alloc_csv) -> Path:
    """Remember what `allocation.csv` looked like when this toolkit last wrote
    or acknowledged it - the baseline `seating.detect_untracked_moves` diffs a
    hand-edited file against."""
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    dest = state_path(cohort)
    shutil.copy2(Path(alloc_csv), dest)
    return dest
