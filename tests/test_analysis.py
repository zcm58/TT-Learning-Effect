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

from LearningEffectAnalysis import extract_trial_number, gather_means_outcome


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
