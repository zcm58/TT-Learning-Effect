#!/usr/bin/env python3
"""
TRIAL ANALYZER APPLICATION

This script provides a graphical user interface (GUI) for analyzing trial-based
experimental data. It supports two main modes of analysis:

1.  Timeline-based Analysis: Compares the first N trials vs. the last N trials
    for a group of participants based on a chronological timeline file.

2.  Outcome-based Analysis: Compares the first N wins vs. the last N wins (or
    losses vs. losses) for a group of participants by sorting the raw data
    files directly, without needing a timeline file.

The application performs the following key operations:
-   Locates and reads experiment data based on the selected analysis mode.
-   Calculates means for each variable for the "first" and "last" trial blocks.
-   Performs appropriate statistical tests (Paired t-test or Wilcoxon signed-rank test).
-   Presents a human-readable summary of significant findings.
-   Allows the full results to be exported to a formatted Excel file.
"""

# =============================================================================
# SECTION 1: DEPENDENCY MANAGEMENT
# =============================================================================
import importlib
import subprocess
import sys
import tkinter
import re

required_packages = [
    'customtkinter',
    'pandas',
    'numpy',
    'scipy',
    'xlsxwriter'
    'openpyxl',  # For reading/writing Excel files
]

for pkg in required_packages:
    try:
        importlib.import_module(pkg)
    except ImportError:
        print(f"Installing missing package: {pkg}")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', pkg])

# =============================================================================
# SECTION 2: IMPORTS
# =============================================================================
import threading
from pathlib import Path
import configparser
import traceback

import customtkinter as ctk
import pandas as pd
import numpy as np
from scipy.stats import shapiro, ttest_rel, wilcoxon
from tkinter import filedialog, messagebox

# =============================================================================
# SECTION 3: CONFIGURATION HELPERS
# =============================================================================
CONFIG_PATH = Path.home() / ".trial_analyzer_config.ini"


def load_default_paths() -> dict:
    """Loads the default directory paths from the config file."""
    config = configparser.ConfigParser()
    paths = {'data_root': '', 'timeline_dir': ''}
    if CONFIG_PATH.exists():
        config.read(CONFIG_PATH)
        if config.has_section("Settings"):
            if config.has_option("Settings", "data_root_dir"):
                data_root = config.get("Settings", "data_root_dir")
                if Path(data_root).is_dir():
                    paths['data_root'] = data_root
            if config.has_option("Settings", "timeline_dir"):
                timeline_dir = config.get("Settings", "timeline_dir")
                if Path(timeline_dir).is_dir():
                    paths['timeline_dir'] = timeline_dir
    return paths


def save_default_paths(data_path: str, timeline_path: str):
    """Saves the given directory paths to the configuration file."""
    config = configparser.ConfigParser()
    config["Settings"] = {
        "data_root_dir": data_path,
        "timeline_dir": timeline_path
    }
    with open(CONFIG_PATH, "w") as f:
        config.write(f)


# =============================================================================
# SECTION 4: CORE ANALYSIS FUNCTIONS
# =============================================================================
# This section is now split into helpers, timeline-based functions, and
# outcome-based functions for clarity.
# -----------------------------------------------------------------------------

# ---[ 4.1 General Helper Functions ]---

def extract_trial_number(path: Path) -> int:
    """Extracts the numerical index from a trial filename."""
    # Uses regex to find the number at the end of the filename, before the extension.
    match = re.search(r'(\d+)\.xls', path.name, re.IGNORECASE)
    return int(match.group(1)) if match else -1


def load_series_from_file(file_path: Path) -> pd.Series:
    """Loads a single trial's data from a direct file path into a pandas Series."""
    df = pd.read_excel(file_path).rename(columns=str.strip)
    if 'Variable' not in df.columns or 'Value' not in df.columns:
        try:
            vcol = next(c for c in df.columns if 'var' in c.lower())
            ycol = next(c for c in df.columns if 'val' in c.lower())
            df = df.rename(columns={vcol: 'Variable', ycol: 'Value'})
        except StopIteration:
            raise ValueError(f"Could not find 'Variable'/'Value' columns in {file_path}")
    return pd.Series(df.Value.values, index=df.Variable.values)


# ---[ 4.2 Timeline-based Analysis Functions ]---

def find_timeline_file(part_dir: Path, timeline_dir: Path, condition: str) -> Path:
    """Finds the correct timeline file for a specific participant and condition."""
    participant_id = part_dir.name.lower()
    condition_lower = condition.lower()
    found_files = []
    if timeline_dir.is_dir():
        for f in timeline_dir.glob('*.xls*'):
            fname_lower = f.name.lower()
            if (fname_lower.startswith(participant_id + "_") and
                    f"_{condition_lower}_" in fname_lower and
                    "timeline" in fname_lower):
                found_files.append(f)
    if not found_files:
        raise FileNotFoundError(f"No timeline file for P '{part_dir.name}'/Cond '{condition}' in {timeline_dir}")
    return found_files[0]


def load_timeline(part_dir: Path, timeline_dir: Path, condition: str) -> list:
    """Loads a timeline file and returns a clean list of trial event IDs."""
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
    """
    Loads a single trial's data based on a constructed trial ID (for Timeline mode).
    This function is flexible to different file naming schemes (e.g., 'P1_Serve_loss1' vs 'P1_Serve_l1')
    and searches for the file within the appropriate 'win' or 'loss' subfolder.
    """
    # 1. Extract outcome string (e.g., 'loss') and number (e.g., '1') from the trial_id
    outcome_str = ''.join(filter(str.isalpha, trial_id))
    number_str = ''.join(filter(str.isdigit, trial_id))
    if not outcome_str or not number_str:
        raise ValueError(f"Invalid trial_id format from timeline: '{trial_id}'")

    # 2. Define the search directory based on the outcome (e.g., .../P11/loss/)
    search_dir = part_dir / outcome_str
    if not search_dir.is_dir():
        raise FileNotFoundError(f"Subfolder '{outcome_str}' not found in {part_dir}")

    # 3. Construct the base of the filename (e.g., 'P11_Serve_')
    base_name = f"{part_dir.name}_{condition}_"

    # 4. Create flexible full filename patterns to search for
    pattern_full = f"{base_name}{outcome_str}{number_str}."  # e.g., 'P11_Serve_loss1.'
    pattern_abbr = f"{base_name}{outcome_str[0]}{number_str}."  # e.g., 'P11_Serve_l1.'

    # 5. Search for a file matching either pattern, case-insensitively
    found_file = None
    for f in search_dir.glob('*.xls*'):
        fname_lower = f.name.lower()
        if fname_lower.startswith(pattern_full.lower()) or fname_lower.startswith(pattern_abbr.lower()):
            found_file = f
            break  # Found a match, stop looking

    if not found_file:
        raise FileNotFoundError(f"No file matching '{pattern_full}' or '{pattern_abbr}' found in {search_dir}")

    return load_series_from_file(found_file)


def gather_means_timeline(part_dir: Path, timeline_dir: Path, condition: str, n: int) -> dict:
    """Calculates means for the first/last N trials based on a timeline."""
    timeline = load_timeline(part_dir, timeline_dir, condition)
    if len(timeline) < 2 * n:
        raise ValueError(f"{part_dir.name} has only {len(timeline)} events (need at least {2 * n})")
    first_ids = timeline[:n]
    last_ids = timeline[-n:]
    first_df = pd.concat([load_trial_series_from_id(part_dir, condition, tid) for tid in first_ids], axis=1)
    last_df = pd.concat([load_trial_series_from_id(part_dir, condition, tid) for tid in last_ids], axis=1)
    return {var: (first_df.loc[var].mean(), last_df.loc[var].mean()) for var in first_df.index}


# ---[ 4.3 Outcome-based Analysis Functions ]---

def gather_means_outcome(part_dir: Path, condition: str, outcome: str, n: int) -> dict:
    """
    Calculates means for the first/last N trials of a specific outcome.
    This function searches for trial files within a subfolder corresponding to the
    outcome (e.g., 'win' or 'loss') inside the participant's directory.
    """
    # Define the search directory based on the outcome (e.g., .../P1/win/)
    outcome_dir = part_dir / outcome.lower()
    if not outcome_dir.is_dir():
        raise FileNotFoundError(f"Outcome folder '{outcome.lower()}' not found for participant {part_dir.name}")

    # Find all Excel files in the outcome directory that have a number in their name.
    all_files = [f for f in outcome_dir.glob('*.xls*') if extract_trial_number(f) != -1]

    # Sort the collected files chronologically based on the extracted trial number.
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


# ---[ 4.4 Main Analysis Runners ]---

def run_analysis(analysis_params: dict, logger=print):
    """
    Executes the entire analysis pipeline across all participants and conditions.
    This function now acts as a router, calling the correct sub-functions based
    on the selected analysis mode.
    """
    results = []
    target_conditions = ['serve', 'return']

    # Unpack parameters from the dictionary
    mode = analysis_params['mode']
    root_dir = analysis_params['data_root']
    n_trials = analysis_params['n_trials']

    for cond_dir in root_dir.iterdir():
        if not cond_dir.is_dir() or cond_dir.name.lower() not in target_conditions:
            if cond_dir.is_dir(): logger(f"Ignoring non-target folder: {cond_dir.name}")
            continue

        condition = cond_dir.name
        agg = {}
        logger(f"Processing Condition: {condition}...")
        for part_dir in cond_dir.iterdir():
            if not part_dir.is_dir(): continue
            try:
                # ROUTER: Call the correct 'gather_means' function based on mode
                if mode == 'timeline':
                    means = gather_means_timeline(
                        part_dir, analysis_params['timeline_dir'], condition, n_trials
                    )
                else:  # mode == 'outcome'
                    means = gather_means_outcome(
                        part_dir, condition, analysis_params['outcome'], n_trials
                    )

                for var, (m1, m2) in means.items():
                    agg.setdefault(var, {'first': [], 'last': []})
                    agg[var]['first'].append(m1)
                    agg[var]['last'].append(m2)

            except (FileNotFoundError, ValueError) as e:
                logger(f"  - Skipping P '{part_dir.name}': {e}")
                continue
            except Exception:
                logger(f"  - An unexpected error occurred with P '{part_dir.name}'. Skipping.")
                logger(traceback.format_exc())
                continue

        # Perform statistical tests on aggregated data for the condition
        for var, data in agg.items():
            a = np.array(data['first'])
            b = np.array(data['last'])
            diffs = a - b
            mean_first, mean_last = np.mean(a), np.mean(b)
            _, sh_p = shapiro(diffs)

            if sh_p > 0.05:
                test_name = 'Paired t-test'
                stat, p = ttest_rel(a, b, nan_policy='omit')
            else:
                test_name = 'Wilcoxon'
                stat, p = (wilcoxon(a, b) if np.any(diffs != 0) else (0, 1))

            # Conditionally create the label based on the analysis mode
            if mode == 'outcome':
                condition_label = f"{condition} ({analysis_params['outcome']})"
            else:  # mode == 'timeline'
                condition_label = condition

            results.append({
                'Condition': condition_label,
                'Variable': var, 'N': len(a), 'Mean_First': mean_first, 'Mean_Last': mean_last,
                'Shapiro_p': sh_p, 'Test': test_name, 'Test_stat': stat, 'p_value': p
            })
    return pd.DataFrame(results)


# =============================================================================
# SECTION 5: GUI APPLICATION
# =============================================================================
class TrialAnalyzerApp(ctk.CTk):
    """The main application window for the Trial Analyzer."""

    def __init__(self):
        super().__init__()
        self.title("Trial Analyzer")
        self.geometry("800x700")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # --- Initialize State Variables ---
        default_paths = load_default_paths()
        self.data_root_var = ctk.StringVar(value=default_paths['data_root'])
        self.timeline_dir_var = ctk.StringVar(value=default_paths['timeline_dir'])
        self.analysis_mode_var = ctk.StringVar(value="timeline")
        self.outcome_var = ctk.StringVar(value="Win")
        self.n_var = ctk.IntVar(value=10)
        self.results_df = None

        # --- Build Menu Bar ---
        menu = tkinter.Menu(self);
        file_menu = tkinter.Menu(menu, tearoff=0)
        file_menu.add_command(label="Exit", command=self.quit)
        menu.add_cascade(label="File", menu=file_menu);
        self.config(menu=menu)

        # --- Build Input Frame ---
        self.input_frame = ctk.CTkFrame(self)
        self.input_frame.pack(fill="x", padx=20, pady=10)

        # Analysis Mode Selection
        mode_frame = ctk.CTkFrame(self.input_frame)
        mode_frame.grid(row=0, column=0, columnspan=4, pady=(5, 10), sticky="w")
        ctk.CTkLabel(mode_frame, text="Analysis Mode:").pack(side="left", padx=5)
        ctk.CTkRadioButton(mode_frame, text="Timeline-based (First N vs Last N)", variable=self.analysis_mode_var,
                           value="timeline", command=self.toggle_mode).pack(side="left", padx=10)
        ctk.CTkRadioButton(mode_frame, text="Outcome-based (e.g., First N Wins vs Last N Wins)",
                           variable=self.analysis_mode_var,
                           value="outcome", command=self.toggle_mode).pack(side="left", padx=10)

        # Data Root Path
        ctk.CTkLabel(self.input_frame, text="Trial Data Root:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ctk.CTkEntry(self.input_frame, textvariable=self.data_root_var, width=450).grid(row=1, column=1, padx=5, pady=5)
        ctk.CTkButton(self.input_frame, text="Browse...", command=self.browse_data_folder).grid(row=1, column=2, padx=5,
                                                                                                pady=5)
        ctk.CTkButton(self.input_frame, text="Save Paths as Default", command=self.save_defaults).grid(row=1, column=3,
                                                                                                       rowspan=2,
                                                                                                       padx=5, pady=5,
                                                                                                       sticky="ns")

        # Timeline-specific Inputs
        self.timeline_label = ctk.CTkLabel(self.input_frame, text="Timeline Files Folder:")
        self.timeline_label.grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.timeline_entry = ctk.CTkEntry(self.input_frame, textvariable=self.timeline_dir_var, width=450)
        self.timeline_entry.grid(row=2, column=1, padx=5, pady=5)
        self.timeline_browse_btn = ctk.CTkButton(self.input_frame, text="Browse...",
                                                 command=self.browse_timeline_folder)
        self.timeline_browse_btn.grid(row=2, column=2, padx=5, pady=5)

        # Outcome-specific Inputs
        self.outcome_label = ctk.CTkLabel(self.input_frame, text="Select Outcome:")
        self.outcome_dropdown = ctk.CTkOptionMenu(self.input_frame, variable=self.outcome_var, values=["Win", "Loss"])

        # Shared Inputs
        ctk.CTkLabel(self.input_frame, text="Trials (first/last N):").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        ctk.CTkEntry(self.input_frame, textvariable=self.n_var, width=80).grid(row=3, column=1, sticky="w", padx=5,
                                                                               pady=5)

        # --- Build Action Buttons Frame ---
        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(fill="x", padx=20)
        self.run_btn = ctk.CTkButton(btn_frame, text="Run Analysis", command=self.start_analysis)
        self.run_btn.pack(side="left", pady=10)
        self.export_btn = ctk.CTkButton(btn_frame, text="Export Results", state="disabled", command=self.export)
        self.export_btn.pack(side="left", padx=10)

        # --- Build Log Textbox ---
        self.log_box = ctk.CTkTextbox(self, width=700, height=350)
        self.log_box.pack(padx=20, pady=10, fill="both", expand=True)

        # Set initial GUI state
        self.toggle_mode()

    def toggle_mode(self):
        """Shows/hides GUI elements based on the selected analysis mode."""
        mode = self.analysis_mode_var.get()
        if mode == "timeline":
            # Show timeline widgets
            self.timeline_label.grid()
            self.timeline_entry.grid()
            self.timeline_browse_btn.grid()
            # Hide outcome widgets
            self.outcome_label.grid_remove()
            self.outcome_dropdown.grid_remove()
        else:  # mode == "outcome"
            # Hide timeline widgets
            self.timeline_label.grid_remove()
            self.timeline_entry.grid_remove()
            self.timeline_browse_btn.grid_remove()
            # Show outcome widgets
            self.outcome_label.grid(row=2, column=0, sticky="w", padx=5, pady=5)
            self.outcome_dropdown.grid(row=2, column=1, sticky="w", padx=5, pady=5)

    def _log(self, message: str):
        self.log_box.insert("end", message + "\n");
        self.log_box.see("end")

    def browse_data_folder(self):
        path = filedialog.askdirectory(title="Select Trial Data Root Folder")
        if path: self.data_root_var.set(path)

    def browse_timeline_folder(self):
        path = filedialog.askdirectory(title="Select Folder Containing All Timeline Files")
        if path: self.timeline_dir_var.set(path)

    def save_defaults(self):
        data_path, timeline_path = self.data_root_var.get(), self.timeline_dir_var.get()
        if not Path(data_path).is_dir():
            messagebox.showerror("Error", "Cannot save: Data Root directory does not exist.")
            return
        save_default_paths(data_path, timeline_path)
        messagebox.showinfo("Saved", "Default paths have been set.")

    def start_analysis(self):
        """Initiates the analysis process in a background thread."""
        self.run_btn.configure(state="disabled")
        self.export_btn.configure(state="disabled")
        self.log_box.delete("1.0", "end")
        self.results_df = None

        # Package all parameters into a dictionary to pass to the thread
        params = {
            'mode': self.analysis_mode_var.get(),
            'data_root': Path(self.data_root_var.get()),
            'n_trials': self.n_var.get(),
            'timeline_dir': Path(self.timeline_dir_var.get()),
            'outcome': self.outcome_var.get()
        }

        threading.Thread(target=self._run_analysis_thread, args=(params,), daemon=True).start()

    def _run_analysis_thread(self, params):
        """The target function that runs the analysis in the background."""
        try:
            if not params['data_root'].is_dir():
                self._log("Error: Data Root directory not found.");
                return
            if params['mode'] == 'timeline' and not params['timeline_dir'].is_dir():
                self._log("Error: Timeline directory not found.");
                return

            self._log("Analysis started...")
            df = run_analysis(params, logger=self._log)
            self.results_df = df

            self._log("\n--- Analysis Complete ---")
            if df.empty:
                self._log("Analysis finished, but no data was successfully processed.")
            else:
                sig = df[df['p_value'] < 0.05]
                if sig.empty:
                    self._log(f"No significant results found at α=0.05 across {len(df)} variables.")
                else:
                    self._log(f"Found {len(sig)} Significant Results (p < 0.05):")
                    for _, row in sig.iterrows():
                        direction = "HIGHER" if row.Mean_Last > row.Mean_First else "LOWER"
                        summary = (f"On average, {row.Variable} was significantly {direction} "
                                   f"in the Last {params['n_trials']} trials.")
                        results_line = (f"Cond: {row.Condition:<20} | Var: {row.Variable:<30} | "
                                        f"Test: {row.Test:<15} | p={row.p_value:.4f}")
                        means_line = (f"  (First {params['n_trials']} Avg: {row.Mean_First:.3f}, "
                                      f"Last {params['n_trials']} Avg: {row.Mean_Last:.3f})")
                        self._log(f"\n{summary}\n{results_line}\n{means_line}")
                self.export_btn.configure(state="normal")
        except Exception:
            self._log("\nAn critical error occurred during analysis.")
            self._log(traceback.format_exc())
        finally:
            self.run_btn.configure(state="normal")

    def export(self):
        """Exports the analysis results DataFrame to a formatted Excel file."""
        if self.results_df is None or self.results_df.empty:
            messagebox.showwarning("No Results", "There are no results to export.");
            return
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Save Results As"
        )
        if not filepath: return
        try:
            with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
                self.results_df.to_excel(writer, sheet_name='Analysis_Results', index=False)
                workbook, worksheet = writer.book, writer.sheets['Analysis_Results']
                center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
                for i, col in enumerate(self.results_df.columns):
                    col_len = max(self.results_df[col].astype(str).map(len).max(), len(col))
                    worksheet.set_column(i, i, col_len + 4, center_format)
            messagebox.showinfo("Success", f"Results successfully exported to\n{filepath}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not save file.\nError: {e}")
            self._log(traceback.format_exc())


# =============================================================================
# SECTION 6: APPLICATION ENTRY POINT
# =============================================================================
if __name__ == "__main__":
    app = TrialAnalyzerApp()
    app.mainloop()

