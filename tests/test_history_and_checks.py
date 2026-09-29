"""Snapshots must round-trip, and drift detection must not invent moves."""
import pandas as pd

import history
import seating


def test_snapshot_and_restore_round_trip(tmp_path, monkeypatch):
    monkeypatch.setattr(history, "HISTORY_DIR", tmp_path / "hist")
    monkeypatch.chdir(tmp_path)
    target = tmp_path / "data" / "file.csv"
    target.parent.mkdir()
    target.write_text("original\n")

    snap = history.snapshot([target.relative_to(tmp_path)], "test")
    assert snap is not None
    target.write_text("clobbered\n")
    history.restore(snap)
    assert target.read_text() == "original\n"


def test_snapshot_of_nothing_is_not_an_error(tmp_path, monkeypatch):
    monkeypatch.setattr(history, "HISTORY_DIR", tmp_path / "hist")
    assert history.snapshot([tmp_path / "does-not-exist.csv"], "test") is None


def test_prune_keeps_the_newest(tmp_path, monkeypatch):
    monkeypatch.setattr(history, "HISTORY_DIR", tmp_path / "hist")
    monkeypatch.chdir(tmp_path)
    f = tmp_path / "f.txt"
    f.write_text("x")
    for i in range(5):
        (tmp_path / "hist").mkdir(exist_ok=True)
        (tmp_path / "hist" / f"2026010{i}-000000_old.zip").write_bytes(b"PK\x05\x06" + b"\0" * 18)
    assert history.prune(keep=2) == 3
    assert len(list((tmp_path / "hist").glob("*.zip"))) == 2


def test_detect_untracked_moves(tmp_path, monkeypatch):
    monkeypatch.setattr(history, "STATE_DIR", tmp_path / "state")
    base = pd.DataFrame({"Roll": ["R1", "R2"], "Name": ["a", "b"], "Block": ["A", "B"],
                         "Seat": ["A-001", "B-001"]})
    live = tmp_path / "allocation.csv"
    base.to_csv(live, index=False)
    history.save_state("english", live)

    assert seating.detect_untracked_moves("english", live).empty

    moved = base.copy()
    moved.loc[moved.Roll == "R1", "Block"] = "C"
    moved.to_csv(live, index=False)
    found = seating.detect_untracked_moves("english", live)
    assert list(found.roll) == ["R1"]
    assert found.iloc[0].from_block == "A" and found.iloc[0].to_block == "C"


def test_no_baseline_reports_nothing_rather_than_everything(tmp_path, monkeypatch):
    monkeypatch.setattr(history, "STATE_DIR", tmp_path / "state")
    live = tmp_path / "allocation.csv"
    pd.DataFrame({"Roll": ["R1"], "Name": ["a"], "Block": ["A"], "Seat": ["A-001"]}).to_csv(live, index=False)
    assert seating.detect_untracked_moves("nosuchcohort", live).empty
