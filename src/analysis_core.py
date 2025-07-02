"""Core analysis utilities for the Trial Analyzer app."""

from pathlib import Path
import re
import configparser
import pandas as pd
import numpy as np
from scipy.stats import shapiro, ttest_rel, wilcoxon

CONFIG_PATH = Path.home() / ".trial_analyzer_config.ini"


def load_default_paths() -> dict:
    """Return default directory paths saved in the config file."""
    config = configparser.ConfigParser()
    paths = {"data_root": "", "timeline_dir": ""}
    if CONFIG_PATH.exists():
        config.read(CONFIG_PATH)
        if config.has_section("Settings"):
            if config.has_option("Settings", "data_root_dir"):
                data_root = config.get("Settings", "data_root_dir")
                if Path(data_root).is_dir():
                    paths["data_root"] = data_root
            if config.has_option("Settings", "timeline_dir"):
                timeline_dir = config.get("Settings", "timeline_dir")
                if Path(timeline_dir).is_dir():
                    paths["timeline_dir"] = timeline_dir
    return paths


def save_default_paths(data_path: str, timeline_path: str) -> None:
    """Persist the given paths to the config file for next launch."""
    config = configparser.ConfigParser()
    config["Settings"] = {"data_root_dir": data_path, "timeline_dir": timeline_path}
    with open(CONFIG_PATH, "w") as f:
        config.write(f)


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def extract_trial_number(path: Path) -> int:
    """Extract and return the trial number encoded in a filename."""
    match = re.search(r"(\d+)\.xls", path.name, re.IGNORECASE)
    return int(match.group(1)) if match else -1


def load_series_from_file(file_path: Path) -> pd.Series:
    """Load a single trial file into a pandas Series."""
    df = pd.read_excel(file_path).rename(columns=str.strip)
    if "Variable" not in df.columns or "Value" not in df.columns:
        try:
            vcol = next(c for c in df.columns if "var" in c.lower())
            ycol = next(c for c in df.columns if "val" in c.lower())
            df = df.rename(columns={vcol: "Variable", ycol: "Value"})
        except StopIteration:
            raise ValueError(f"Could not find 'Variable'/'Value' columns in {file_path}")
    return pd.Series(df.Value.values, index=df.Variable.values)


def find_timeline_file(part_dir: Path, timeline_dir: Path, condition: str) -> Path:
    """Locate the timeline file for a participant and condition."""
    participant_id = part_dir.name.lower()
    condition_lower = condition.lower()
    found_files = []
    if timeline_dir.is_dir():
        for f in timeline_dir.glob("*.xls*"):
            fname_lower = f.name.lower()
            if (
                fname_lower.startswith(participant_id + "_")
                and f"_{condition_lower}_" in fname_lower
                and "timeline" in fname_lower
            ):
                found_files.append(f)
    if not found_files:
        raise FileNotFoundError(
            f"No timeline file for P '{part_dir.name}'/Cond '{condition}' in {timeline_dir}"
        )
    return found_files[0]


def load_timeline(part_dir: Path, timeline_dir: Path, condition: str) -> list:
    """Return the ordered list of trial IDs from a participant's timeline file."""
    timeline_file_path = find_timeline_file(part_dir, timeline_dir, condition)
    df = pd.read_excel(timeline_file_path)
    try:
        type_col = next(c for c in df.columns if "type" in c.lower())
        index_col = next(c for c in df.columns if "trial" in c.lower())
    except StopIteration:
        raise ValueError("Could not find 'type' and 'trial' columns in timeline file.")
    events = df[type_col].astype(str).str.strip().str.lower().tolist()
    idxs = df[index_col].astype(str).str.strip().tolist()
    return [f"{ev}{ix}" for ev, ix in zip(events, idxs)]


def load_trial_series_from_id(part_dir: Path, condition: str, trial_id: str) -> pd.Series:
    """Load a single trial identified by a timeline event ID."""
    outcome_str = "".join(filter(str.isalpha, trial_id))
    number_str = "".join(filter(str.isdigit, trial_id))
    if not outcome_str or not number_str:
        raise ValueError(f"Invalid trial_id format from timeline: '{trial_id}'")

    search_dir = part_dir / outcome_str
    if not search_dir.is_dir():
        raise FileNotFoundError(f"Subfolder '{outcome_str}' not found in {part_dir}")

    base_name = f"{part_dir.name}_{condition}_"
    pattern_full = f"{base_name}{outcome_str}{number_str}."
    pattern_abbr = f"{base_name}{outcome_str[0]}{number_str}."

    found_file = None
    for f in search_dir.glob("*.xls*"):
        fname_lower = f.name.lower()
        if fname_lower.startswith(pattern_full.lower()) or fname_lower.startswith(
            pattern_abbr.lower()
        ):
            found_file = f
            break
    if not found_file:
        raise FileNotFoundError(
            f"No file matching '{pattern_full}' or '{pattern_abbr}' found in {search_dir}"
        )
    return load_series_from_file(found_file)


def gather_means_timeline(part_dir: Path, timeline_dir: Path, condition: str, n: int) -> dict:
    """Compute per-variable means for the first and last N timeline events."""
    timeline = load_timeline(part_dir, timeline_dir, condition)
    if len(timeline) < 2 * n:
        raise ValueError(f"{part_dir.name} has only {len(timeline)} events (need at least {2 * n})")
    first_ids = timeline[:n]
    last_ids = timeline[-n:]
    first_df = pd.concat([load_trial_series_from_id(part_dir, condition, tid) for tid in first_ids], axis=1)
    last_df = pd.concat([load_trial_series_from_id(part_dir, condition, tid) for tid in last_ids], axis=1)
    return {var: (first_df.loc[var].mean(), last_df.loc[var].mean()) for var in first_df.index}


def gather_means_outcome(part_dir: Path, condition: str, outcome: str, n: int) -> dict:
    """Compute means for the first/last N trials within an outcome subfolder."""
    outcome_dir = part_dir / outcome.lower()
    if not outcome_dir.is_dir():
        raise FileNotFoundError(
            f"Outcome folder '{outcome.lower()}' not found for participant {part_dir.name}"
        )
    all_files = [f for f in outcome_dir.glob("*.xls*") if extract_trial_number(f) != -1]
    files = sorted(all_files, key=extract_trial_number)
    if not files:
        raise FileNotFoundError(f"No valid trial files found in '{outcome_dir}'")
    if len(files) < 2 * n:
        raise ValueError(f"{part_dir.name} has only {len(files)} '{outcome}' files (need at least {2 * n})")

    first_files = files[:n]
    last_files = files[-n:]
    first_df = pd.concat([load_series_from_file(f) for f in first_files], axis=1)
    last_df = pd.concat([load_series_from_file(f) for f in last_files], axis=1)
    return {var: (first_df.loc[var].mean(), last_df.loc[var].mean()) for var in first_df.index}


def run_analysis(analysis_params: dict, logger=print) -> pd.DataFrame:
    """Execute the analysis pipeline across all participants and conditions."""
    results = []
    target_conditions = ["serve", "return"]

    mode = analysis_params["mode"]
    root_dir = analysis_params["data_root"]
    n_trials = analysis_params["n_trials"]

    for cond_dir in root_dir.iterdir():
        if not cond_dir.is_dir() or cond_dir.name.lower() not in target_conditions:
            if cond_dir.is_dir():
                logger(f"Ignoring non-target folder: {cond_dir.name}")
            continue

        condition = cond_dir.name
        agg: dict[str, dict[str, list]] = {}
        logger(f"Processing Condition: {condition}...")
        for part_dir in cond_dir.iterdir():
            if not part_dir.is_dir():
                continue
            try:
                if mode == "timeline":
                    means = gather_means_timeline(part_dir, analysis_params["timeline_dir"], condition, n_trials)
                else:
                    means = gather_means_outcome(part_dir, condition, analysis_params["outcome"], n_trials)

                for var, (m1, m2) in means.items():
                    agg.setdefault(var, {"first": [], "last": []})
                    agg[var]["first"].append(m1)
                    agg[var]["last"].append(m2)

            except (FileNotFoundError, ValueError) as e:
                logger(f"  - Skipping P '{part_dir.name}': {e}")
                continue
            except Exception:
                logger(f"  - An unexpected error occurred with P '{part_dir.name}'. Skipping.")
                logger(traceback.format_exc())
                continue

        for var, data in agg.items():
            a = np.array(data["first"])
            b = np.array(data["last"])
            diffs = a - b
            mean_first, mean_last = np.mean(a), np.mean(b)
            _, sh_p = shapiro(diffs)
            if sh_p > 0.05:
                test_name = "Paired t-test"
                stat, p = ttest_rel(a, b, nan_policy="omit")
            else:
                test_name = "Wilcoxon"
                stat, p = (wilcoxon(a, b) if np.any(diffs != 0) else (0, 1))

            if mode == "outcome":
                condition_label = f"{condition} ({analysis_params['outcome']})"
            else:
                condition_label = condition

            results.append(
                {
                    "Condition": condition_label,
                    "Variable": var,
                    "N": len(a),
                    "Mean_First": mean_first,
                    "Mean_Last": mean_last,
                    "Shapiro_p": sh_p,
                    "Test": test_name,
                    "Test_stat": stat,
                    "p_value": p,
                }
            )
    return pd.DataFrame(results)

