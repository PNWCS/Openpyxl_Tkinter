"""GUI module for XLSX Reader application.

Provides a Tkinter interface to select an Excel file (.xlsx),
show live progress while processing, and display per-sheet row counts.
"""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk
from typing import Literal

from .excel_processor import process_excel_file


def select_excel_file() -> str:
    """Open a file dialog to select an Excel file (.xlsx)."""
    path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Workbook", "*.xlsx")],
        defaultextension=".xlsx",
    )
    return path or ""


def update_progress(progress_bar: ttk.Progressbar, current: int, total: int) -> None:
    """Update the determinate progress bar using a zero-based `current` index."""
    if total <= 0:
        progress_bar["value"] = 0
        return
    percent = int(((current + 1) / total) * 100)
    if percent < 0:
        percent = 0
    if percent > 100:
        percent = 100
    # subscript style keeps mypy happy with tk stubs
    progress_bar["value"] = percent


def process_file_in_background(
    file_path: str,
    progress_bar: ttk.Progressbar,
    status_label: tk.Label,
    process_button: tk.Button,
    results_text: tk.Text,
) -> None:
    """Run processing in a worker thread and marshal UI updates onto the main thread."""

    # --- UI helpers (schedule back to main thread) ---
    def ui_set_status(text: str) -> None:
        status_label.after(0, lambda: status_label.config(text=text))

    def ui_set_button_state(state: Literal["normal", "disabled", "active"]) -> None:
        process_button.after(0, lambda: process_button.config(state=state))

    def ui_set_progress(current: int, total: int) -> None:
        progress_bar.after(0, lambda: update_progress(progress_bar, current, total))

    def ui_set_progress_value(value: int) -> None:
        progress_bar.after(0, lambda: progress_bar.__setitem__("value", value))

    def ui_show_results(results: dict[str, int]) -> None:
        def _render() -> None:
            results_text.delete("1.0", tk.END)
            results_text.insert(tk.END, "Row counts per sheet:\n")
            results_text.insert(tk.END, "-" * 30 + "\n")
            for sheet_name, row_count in results.items():
                results_text.insert(tk.END, f"{sheet_name}: {row_count} rows\n")
            total_rows = sum(results.values())
            results_text.insert(tk.END, "-" * 30 + "\n")
            results_text.insert(tk.END, f"Total rows: {total_rows}\n")
            results_text.insert(tk.END, f"Total sheets: {len(results)}\n")

        results_text.after(0, _render)

    def ui_show_error(msg: str) -> None:
        def _render() -> None:
            results_text.delete("1.0", tk.END)
            results_text.insert(tk.END, f"Error processing file:\n{msg}")

        results_text.after(0, _render)

    # --- Worker function ---
    def worker() -> None:
        try:
            ui_set_button_state("disabled")
            ui_set_status("Processing...")
            ui_set_progress_value(0)

            def progress_callback(current: int, total: int, sheet_name: str) -> None:
                ui_set_status(f"Processing sheet: {sheet_name}")
                ui_set_progress(current, total)

            results = process_excel_file(file_path, progress_callback)

            ui_set_progress_value(100)
            ui_set_status("Processing complete!")
            ui_show_results(results)

        except Exception as exc:  # noqa: BLE001
            ui_set_status("Error occurred!")
            ui_show_error(str(exc))
        finally:
            ui_set_button_state("normal")

    if not Path(file_path).exists():
        ui_set_status("Selected file not found.")
        return

    threading.Thread(target=worker, daemon=True).start()


def create_main_window() -> tk.Tk:
    """Create and configure the main application window."""
    root = tk.Tk()
    root.title("XLSX Reader")
    root.geometry("500x400")
    root.resizable(True, True)
    return root


def run_app() -> None:
    """Build the UI and start the Tkinter main loop."""
    root = create_main_window()

    # Top controls
    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    select_button = tk.Button(
        button_frame,
        text="Select Excel File",
        font=("Arial", 12),
        bg="lightblue",
        width=20,
    )
    select_button.pack()

    # Progress
    progress_frame = tk.Frame(root)
    progress_frame.pack(pady=5)

    progress_bar = ttk.Progressbar(progress_frame, length=400, mode="determinate", maximum=100)
    progress_bar.pack()

    # Status
    status_frame = tk.Frame(root)
    status_frame.pack(pady=5)

    status_label = tk.Label(status_frame, text="Select an Excel file to begin", font=("Arial", 10))
    status_label.pack()

    # Results (with scrollbar)
    results_frame = tk.Frame(root)
    results_frame.pack(pady=10, padx=20, fill="both", expand=True)

    results_label = tk.Label(results_frame, text="Results:", font=("Arial", 10, "bold"))
    results_label.pack(anchor="w")

    text_frame = tk.Frame(results_frame)
    text_frame.pack(fill="both", expand=True)

    results_text = tk.Text(text_frame, height=12, width=50, font=("Courier", 10), wrap=tk.WORD)
    scrollbar = tk.Scrollbar(text_frame, orient="vertical", command=results_text.yview)
    results_text.configure(yscrollcommand=scrollbar.set)

    results_text.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Button action
    def on_select_file() -> None:
        path = select_excel_file()
        if path:
            process_file_in_background(
                path, progress_bar, status_label, select_button, results_text
            )

    select_button.config(command=on_select_file)

    root.mainloop()


if __name__ == "__main__":
    run_app()
