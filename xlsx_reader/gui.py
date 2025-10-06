from __future__ import annotations

import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, Optional

from .excel_processor import process_excel_file


def select_excel_file() -> str:
    """Open a file picker and return the selected Excel file path (or '')."""
    path = filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=[("Excel files", "*.xlsx;*.xls")],
    )
    return path or ""


def update_progress(progressbar: ttk.Progressbar, var: tk.DoubleVar, current: int, total: int) -> None:
    """Update progressbar (0..1) given current/total steps."""
    total = max(1, int(total))
    value = max(0.0, min(1.0, float(current) / float(total)))
    var.set(value)
    progressbar.update_idletasks()


def process_file_in_background(
    file_path: str,
    progressbar: ttk.Progressbar,
    progress_var: tk.DoubleVar,
    output_text: tk.Text,
    start_button: tk.Button,
    status_label: ttk.Label,
) -> None:
    """
    Spawn a worker thread that processes the Excel file so the UI stays responsive.
    Uses process_excel_file's callback signature: (current, total, sheet_name).
    """

    def worker() -> None:
        try:
            def on_progress(current: int, total: int, sheet: str) -> None:
                # marshal back to UI thread
                progressbar.after(0, update_progress, progressbar, progress_var, current, total)
                status_label.after(0, lambda: status_label.config(text=f"Processing: {sheet} ({current}/{total})"))

            results: Dict[str, int] = process_excel_file(file_path, progress_callback=on_progress)

            def write_results() -> None:
                output_text.config(state="normal")
                output_text.delete("1.0", tk.END)
                if not results:
                    output_text.insert(tk.END, "No sheets found.\n")
                else:
                    for sheet, count in results.items():
                        output_text.insert(tk.END, f"{sheet}: {count} rows\n")
                output_text.config(state="disabled")
                status_label.config(text="Done.")
                start_button.config(state="normal")

            output_text.after(0, write_results)

        except FileNotFoundError:
            def show_missing() -> None:
                messagebox.showerror("File not found", "The selected file could not be found.")
                status_label.config(text="Ready.")
                start_button.config(state="normal")
            progressbar.after(0, show_missing)

        except Exception as e:
            def show_err() -> None:
                messagebox.showerror("Error", str(e))
                status_label.config(text="Ready.")
                start_button.config(state="normal")
            progressbar.after(0, show_err)

    # disable button while processing
    start_button.config(state="disabled")
    status_label.config(text="Starting…")
    threading.Thread(target=worker, daemon=True).start()


def create_main_window() -> tk.Tk:
    root = tk.Tk()
    root.title("XLSX Reader")
    root.geometry("700x420")

    # Main frame
    frm = ttk.Frame(root, padding=12)
    frm.grid(sticky="nsew")
    root.rowconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)

    # --- File chooser row ---
    file_var = tk.StringVar(value="")
    ttk.Label(frm, text="Excel file:").grid(row=0, column=0, sticky="w", padx=(0, 8))

    file_entry = ttk.Entry(frm, textvariable=file_var, width=60)
    file_entry.grid(row=0, column=1, sticky="ew")
    frm.columnconfigure(1, weight=1)

    def on_select_file() -> None:
        path = select_excel_file()
        if path:
            file_var.set(path)

    browse_btn = ttk.Button(frm, text="Browse…", command=on_select_file)
    browse_btn.grid(row=0, column=2, padx=(8, 0))

    # --- Progress + status ---
    progress_var = tk.DoubleVar(value=0.0)
    progress = ttk.Progressbar(frm, orient="horizontal", mode="determinate",
                               variable=progress_var, maximum=1.0, length=200)
    progress.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(12, 6))

    status_label = ttk.Label(frm, text="Ready.")
    status_label.grid(row=2, column=0, columnspan=3, sticky="w", pady=(0, 6))

    # --- Output text ---
    output = tk.Text(frm, height=14, width=80, state="disabled")
    output.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=(4, 0))
    frm.rowconfigure(3, weight=1)

    # --- Buttons row ---
    def on_start() -> None:
        path = file_var.get().strip()
        if not path:
            messagebox.showwarning("Missing file", "Please choose an Excel file first.")
            return
        # reset UI
        progress_var.set(0.0)
        status_label.config(text="Queued…")
        output.config(state="normal")
        output.delete("1.0", tk.END)
        output.config(state="disabled")

        process_file_in_background(
            path,
            progressbar=progress,
            progress_var=progress_var,
            output_text=output,
            start_button=start_btn,
            status_label=status_label,
        )

    start_btn = ttk.Button(frm, text="Process", command=on_start)
    start_btn.grid(row=4, column=0, sticky="w", pady=(10, 0))

    quit_btn = ttk.Button(frm, text="Quit", command=root.destroy)
    quit_btn.grid(row=4, column=2, sticky="e", pady=(10, 0))

    return root


def run_app() -> None:
    root = create_main_window()
    root.mainloop()
