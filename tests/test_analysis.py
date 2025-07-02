import sys
import types
from pathlib import Path
import pytest

class DummyModule(types.ModuleType):
    def __getattr__(self, name):
        class _Dummy:
            def __init__(self, *a, **k):
                pass
        return _Dummy

# Provide dummy modules for optional dependencies to avoid pip installs
for mod in ["customtkinter", "openpyxl", "xlsxwriter", "xlsxwriteropenpyxl"]:
    if mod not in sys.modules:
        sys.modules[mod] = DummyModule(mod)

# Ensure the src directory is on the Python path
sys.path.append(str(Path(__file__).resolve().parents[1] / 'src'))

import pandas as pd
from LearningEffectAnalysis import (
    extract_trial_number,
    gather_means_outcome,
    gather_means_timeline,
    run_analysis,
)


@pytest.mark.parametrize(
    "filename,expected",
    [
        ("P1_serve_win1.xls", 1),
        ("P2_return_L10.xlsx", 10),
        ("trial07.xlsm", 7),
        ("random_file.txt", -1),
    ],
)
def test_extract_trial_number_parsing(filename, expected):
    assert extract_trial_number(Path(filename)) == expected


def test_gather_means_outcome_missing_dir(tmp_path):
    part_dir = tmp_path / "P1"
    part_dir.mkdir()
    with pytest.raises(FileNotFoundError):
        gather_means_outcome(part_dir, "Serve", "Win", 5)


def test_gather_means_outcome_no_files(tmp_path):
    part_dir = tmp_path / "P1"
    outcome_dir = part_dir / "win"
    outcome_dir.mkdir(parents=True)
    with pytest.raises(FileNotFoundError):
        gather_means_outcome(part_dir, "Serve", "Win", 5)


def test_gather_means_outcome_insufficient_files(tmp_path):
    part_dir = tmp_path / "P1"
    outcome_dir = part_dir / "win"
    outcome_dir.mkdir(parents=True)
    for i in range(3):
        (outcome_dir / f"P1_Serve_win{i+1}.xls").touch()
    with pytest.raises(ValueError):
        gather_means_outcome(part_dir, "Serve", "Win", 2)


def test_gather_means_outcome_success(tmp_path, monkeypatch):
    part_dir = tmp_path / "P1"
    outcome_dir = part_dir / "win"
    outcome_dir.mkdir(parents=True)
    for i in range(4):
        df = pd.DataFrame({"Variable": ["A", "B"], "Value": [i + 1, (i + 1) * 10]})
        df.to_csv(outcome_dir / f"P1_Serve_win{i+1}.xls", index=False)

    monkeypatch.setattr("pandas.read_excel", lambda path, *a, **k: pd.read_csv(path))

    means = gather_means_outcome(part_dir, "Serve", "Win", 2)
    assert means["A"] == (1.5, 3.5)
    assert means["B"] == (15.0, 35.0)


def test_gather_means_timeline_success(tmp_path, monkeypatch):
    part_dir = tmp_path / "P1"
    win_dir = part_dir / "win"
    win_dir.mkdir(parents=True)
    for i in range(4):
        df = pd.DataFrame({"Variable": ["A", "B"], "Value": [i + 1, (i + 1) * 10]})
        df.to_csv(win_dir / f"P1_Serve_win{i+1}.xls", index=False)

    timeline_dir = tmp_path / "timelines"
    timeline_dir.mkdir()
    tl_df = pd.DataFrame({"Type": ["win"] * 4, "Trial": [1, 2, 3, 4]})
    tl_df.to_csv(timeline_dir / "P1_Serve_timeline.xls", index=False)

    monkeypatch.setattr("pandas.read_excel", lambda path, *a, **k: pd.read_csv(path))

    means = gather_means_timeline(part_dir, timeline_dir, "Serve", 2)
    assert means["A"] == (1.5, 3.5)
    assert means["B"] == (15.0, 35.0)


def test_run_analysis_end_to_end(tmp_path, monkeypatch):
    data_root = tmp_path / "data"
    cond_dir = data_root / "serve"
    participants = ["P1", "P2", "P3"]
    for pid in participants:
        p_dir = cond_dir / pid / "win"
        p_dir.mkdir(parents=True, exist_ok=True)
        for i in range(4):
            df = pd.DataFrame({"Variable": ["A", "B"], "Value": [i + 1, (i + 1) * 10]})
            df.to_csv(p_dir / f"{pid}_serve_win{i+1}.xls", index=False)

    monkeypatch.setattr("pandas.read_excel", lambda path, *a, **k: pd.read_csv(path))

    params = {"mode": "outcome", "data_root": data_root, "outcome": "win", "n_trials": 2}
    df = run_analysis(params, logger=lambda *a, **k: None)

    assert not df.empty
    assert set(df.Variable) == {"A", "B"}
    assert df.N.nunique() == 1 and df.N.iloc[0] == len(participants)
    row_a = df[df.Variable == "A"].iloc[0]
    assert row_a.Mean_First == pytest.approx(1.5)
    assert row_a.Mean_Last == pytest.approx(3.5)
