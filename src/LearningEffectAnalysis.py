#!/usr/bin/env python3
"""GUI application for analyzing learning effects in trial-based data."""

import threading
import tkinter
from pathlib import Path
import traceback

import customtkinter as ctk
import pandas as pd
from tkinter import filedialog, messagebox

from analysis_core import (
    load_default_paths,
    save_default_paths,
    extract_trial_number,
    load_series_from_file,
    find_timeline_file,
    load_timeline,
    load_trial_series_from_id,
    gather_means_timeline,
    gather_means_outcome,
    run_analysis,
)
import gui_animations as anim


class TrialAnalyzerApp(ctk.CTk):
    """Main application window for the Trial Analyzer."""

    def __init__(self) -> None:
        """Initialize the GUI components and application state."""
        super().__init__()
        self.title("Trial Analyzer")
        self.geometry("800x700")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # --- Initialize State Variables ---
        default_paths = load_default_paths()
        self.data_root_var = ctk.StringVar(value=default_paths["data_root"])
        self.timeline_dir_var = ctk.StringVar(value=default_paths["timeline_dir"])
        self.analysis_mode_var = ctk.StringVar(value="timeline")
        self.outcome_var = ctk.StringVar(value="Win")
        self.n_var = ctk.IntVar(value=10)
        self.results_df = None

        # --- Build Menu Bar ---
        menu = tkinter.Menu(self)
        file_menu = tkinter.Menu(menu, tearoff=0)
        file_menu.add_command(label="Exit", command=self.quit)
        menu.add_cascade(label="File", menu=file_menu)
        self.config(menu=menu)

        # --- Build Input Frame ---
        self.input_frame = ctk.CTkFrame(self)
        self.input_frame.pack(fill="x", padx=20, pady=10)

        mode_frame = ctk.CTkFrame(self.input_frame)
        mode_frame.grid(row=0, column=0, columnspan=4, pady=(5, 10), sticky="w")
        ctk.CTkLabel(mode_frame, text="Analysis Mode:").pack(side="left", padx=5)
        ctk.CTkRadioButton(mode_frame, text="Timeline-based (First N vs Last N)",
                           variable=self.analysis_mode_var,
                           value="timeline", command=self.toggle_mode).pack(side="left", padx=10)
        ctk.CTkRadioButton(mode_frame, text="Outcome-based (e.g., First N Wins vs Last N Wins)",
                           variable=self.analysis_mode_var,
                           value="outcome", command=self.toggle_mode).pack(side="left", padx=10)

        ctk.CTkLabel(self.input_frame, text="Trial Data Root:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ctk.CTkEntry(self.input_frame, textvariable=self.data_root_var, width=450).grid(row=1, column=1, padx=5, pady=5)
        ctk.CTkButton(self.input_frame, text="Browse...", command=self.browse_data_folder).grid(row=1, column=2, padx=5, pady=5)
        ctk.CTkButton(self.input_frame, text="Save Paths as Default", command=self.save_defaults).grid(row=1, column=3, rowspan=2, padx=5, pady=5, sticky="ns")

        self.timeline_label = ctk.CTkLabel(self.input_frame, text="Timeline Files Folder:")
        self.timeline_label.grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.timeline_entry = ctk.CTkEntry(self.input_frame, textvariable=self.timeline_dir_var, width=450)
        self.timeline_entry.grid(row=2, column=1, padx=5, pady=5)
        self.timeline_browse_btn = ctk.CTkButton(self.input_frame, text="Browse...", command=self.browse_timeline_folder)
        self.timeline_browse_btn.grid(row=2, column=2, padx=5, pady=5)

        self.outcome_label = ctk.CTkLabel(self.input_frame, text="Select Outcome:")
        self.outcome_dropdown = ctk.CTkOptionMenu(self.input_frame, variable=self.outcome_var, values=["Win", "Loss"])

        ctk.CTkLabel(self.input_frame, text="Trials (first/last N):").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        ctk.CTkEntry(self.input_frame, textvariable=self.n_var, width=80).grid(row=3, column=1, sticky="w", padx=5, pady=5)

        # --- Build Action Buttons Frame ---
        self.btn_frame = ctk.CTkFrame(self)
        self.btn_frame.pack(fill="x", padx=20)
        self.run_btn = ctk.CTkButton(self.btn_frame, text="Run Analysis", command=self.start_analysis)
        self.run_btn.pack(side="left", pady=10)
        self.export_btn = ctk.CTkButton(self.btn_frame, text="Export Results", state="disabled", command=self.export)
        self.export_btn.pack(side="left", padx=10)
        self.progress = ctk.CTkProgressBar(self.btn_frame, mode="indeterminate")
        self.progress.pack(side="left", fill="x", expand=True, padx=10)
        self.progress.stop()
        self.progress.pack_forget()

        # --- Build Log Textbox ---
        self.log_box = ctk.CTkTextbox(self, width=700, height=350)
        self.log_box.pack(padx=20, pady=10, fill="both", expand=True)

        self.toggle_mode()


    # ------------------------------------------------------------------
    # GUI helper methods
    # ------------------------------------------------------------------
    def toggle_mode(self) -> None:
        """Animate switching between analysis modes."""
        mode = self.analysis_mode_var.get()
        def update_widgets() -> None:
            if mode == "timeline":
                self.timeline_label.grid()
                self.timeline_entry.grid()
                self.timeline_browse_btn.grid()
                self.outcome_label.grid_remove()
                self.outcome_dropdown.grid_remove()
            else:
                self.timeline_label.grid_remove()
                self.timeline_entry.grid_remove()
                self.timeline_browse_btn.grid_remove()
                self.outcome_label.grid(row=2, column=0, sticky="w", padx=5, pady=5)
                self.outcome_dropdown.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        anim.fade_window(self, update_widgets)

    def _log(self, message: str) -> None:
        """Append a message to the log textbox."""
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")

    def browse_data_folder(self) -> None:
        """Prompt the user to select the root directory containing trial data."""
        path = filedialog.askdirectory(title="Select Trial Data Root Folder")
        if path:
            self.data_root_var.set(path)

    def browse_timeline_folder(self) -> None:
        """Prompt the user to choose the folder with all timeline files."""
        path = filedialog.askdirectory(title="Select Folder Containing All Timeline Files")
        if path:
            self.timeline_dir_var.set(path)

    def save_defaults(self) -> None:
        """Persist the currently selected paths as defaults."""
        data_path = self.data_root_var.get()
        timeline_path = self.timeline_dir_var.get()
        if not Path(data_path).is_dir():
            messagebox.showerror("Error", "Cannot save: Data Root directory does not exist.")
            return
        save_default_paths(data_path, timeline_path)
        anim.popup(self, "Saved", "Default paths have been set.")

    def start_analysis(self) -> None:
        """Initiate the analysis in a background thread."""
        anim.pulse(self.run_btn)
        self.run_btn.configure(state="disabled")
        self.export_btn.configure(state="disabled")
        self.log_box.delete("1.0", "end")
        self.results_df = None

        params = {
            "mode": self.analysis_mode_var.get(),
            "data_root": Path(self.data_root_var.get()),
            "n_trials": self.n_var.get(),
            "timeline_dir": Path(self.timeline_dir_var.get()),
            "outcome": self.outcome_var.get(),
        }
        self.progress.pack(side="left", fill="x", expand=True, padx=10)
        self.progress.start()
        threading.Thread(target=self._run_analysis_thread, args=(params,), daemon=True).start()

    def _run_analysis_thread(self, params: dict) -> None:
        """Run the analysis and update the UI when done."""
        try:
            if not params["data_root"].is_dir():
                self._log("Error: Data Root directory not found.")
                return
            if params["mode"] == "timeline" and not params["timeline_dir"].is_dir():
                self._log("Error: Timeline directory not found.")
                return
            self._log("Analysis started...")
            df = run_analysis(params, logger=self._log)
            self.results_df = df
            self._log("\n--- Analysis Complete ---")
            if df.empty:
                self._log("Analysis finished, but no data was successfully processed.")
            else:
                sig = df[df["p_value"] < 0.05]
                if sig.empty:
                    self._log(f"No significant results found at α=0.05 across {len(df)} variables.")
                else:
                    self._log(f"Found {len(sig)} Significant Results (p < 0.05):")
                    for _, row in sig.iterrows():
                        direction = "HIGHER" if row.Mean_Last > row.Mean_First else "LOWER"
                        summary = (
                            f"On average, {row.Variable} was significantly {direction} "
                            f"in the Last {params['n_trials']} trials."
                        )
                        results_line = (
                            f"Cond: {row.Condition:<20} | Var: {row.Variable:<30} | "
                            f"Test: {row.Test:<15} | p={row.p_value:.4f}"
                        )
                        means_line = (
                            f"  (First {params['n_trials']} Avg: {row.Mean_First:.3f}, "
                            f"Last {params['n_trials']} Avg: {row.Mean_Last:.3f})"
                        )
                        self._log(f"\n{summary}\n{results_line}\n{means_line}")
                self.export_btn.configure(state="normal")
        except Exception:
            self._log("\nA critical error occurred during analysis.")
            self._log(traceback.format_exc())
        finally:
            self.progress.stop()
            self.progress.pack_forget()
            self.run_btn.configure(state="normal")

    def export(self) -> None:
        """Export the results DataFrame to Excel."""
        if self.results_df is None or self.results_df.empty:
            messagebox.showwarning("No Results", "There are no results to export.")
            return
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Save Results As",
        )
        if not filepath:
            return
        try:
            with pd.ExcelWriter(filepath, engine="xlsxwriter") as writer:
                self.results_df.to_excel(writer, sheet_name="Analysis_Results", index=False)
                workbook, worksheet = writer.book, writer.sheets["Analysis_Results"]
                center_format = workbook.add_format({"align": "center", "valign": "vcenter"})
                for i, col in enumerate(self.results_df.columns):
                    col_len = max(self.results_df[col].astype(str).map(len).max(), len(col))
                    worksheet.set_column(i, i, col_len + 4, center_format)
            anim.pulse(self.export_btn)
            anim.popup(self, "Success", f"Results successfully exported to\n{filepath}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not save file.\nError: {e}")
            self._log(traceback.format_exc())


if __name__ == "__main__":
    app = TrialAnalyzerApp()
    app.mainloop()
