#!/usr/bin/env python3
"""Advisory "someone else has this open" marker for the Dropbox-synced folder.

Two machines running the dashboard at once is the one failure this project
warns about in prose but has never actually detected: Dropbox syncs on save,
not mid-write, and a regeneration rewrites a dozen files in quick succession,
so two devices writing around the same moment leaves a `(Conflicted copy)`
file that then has to be untangled by hand.

This can't *prevent* that - Dropbox propagation is seconds to minutes, and a
lock that arrives late is no lock at all. What it can do is tell you, when you
open the dashboard, that another machine was active a few minutes ago, which
is the moment the warning is actually useful. Deliberately advisory: it never
blocks, because a stale marker that locks you out of your own tool is worse
than the conflict it was guarding against.
"""
from __future__ import annotations

import json
import os
import socket
from datetime import datetime, timedelta
from pathlib import Path

LOCK_PATH = Path(".dashboard.lock")
STALE_AFTER = timedelta(minutes=5)


def _me() -> dict:
    return {"host": socket.gethostname(), "pid": os.getpid()}


def _read(path: Path) -> list[dict]:
    if not path.exists():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except (json.JSONDecodeError, OSError):
        return []          # a half-synced file is not worth an exception


def heartbeat(path: str | Path = LOCK_PATH) -> list[dict]:
    """Record that this process is alive; return *other* live sessions.

    Entries older than `STALE_AFTER` are dropped, so a machine that was shut
    down without cleaning up stops being reported within minutes.
    """
    path = Path(path)
    me = _me()
    now = datetime.now()
    keep = []
    others = []
    for e in _read(path):
        if e.get("host") == me["host"] and e.get("pid") == me["pid"]:
            continue
        try:
            seen = datetime.fromisoformat(e.get("heartbeat", ""))
        except ValueError:
            continue
        if now - seen > STALE_AFTER:
            continue
        keep.append(e)
        others.append({**e, "minutes": round((now - seen).total_seconds() / 60, 1)})

    keep.append({**me, "heartbeat": now.isoformat(timespec="seconds")})
    try:
        path.write_text(json.dumps(keep, indent=1), encoding="utf-8")
    except OSError:
        pass               # read-only folder shouldn't take the dashboard down
    return others


def release(path: str | Path = LOCK_PATH) -> None:
    path = Path(path)
    me = _me()
    keep = [e for e in _read(path)
            if not (e.get("host") == me["host"] and e.get("pid") == me["pid"])]
    try:
        path.write_text(json.dumps(keep, indent=1), encoding="utf-8") if keep else path.unlink()
    except OSError:
        pass


def describe(others: list[dict]) -> str:
    if not others:
        return ""
    who = ", ".join(f"**{o['host']}** (last seen {o['minutes']} min ago)" for o in others[:3])
    return (f"Another dashboard is open on {who}. This folder syncs through Dropbox, and two "
            "machines regenerating files at the same time produces a *(Conflicted copy)* file "
            "that has to be untangled by hand. Safe to read here; avoid building or moving "
            "anyone until the other one is closed.")
