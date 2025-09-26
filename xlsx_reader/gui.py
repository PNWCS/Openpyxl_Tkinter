"""GUI module for XLSX Reader application.

This module provides a simple tkinter interface for selecting Excel files
and showing progress while processing them.
"""

import threading
import tkinter as tk
from tkinter import filedialog, ttk

from .excel_processor import process_excel_file


def select_excel_file() -> str:
    """Open a file dialog to select an Excel file.

    Returns:
        str: Path to selected Excel file, or empty string if cancelled

    Note:
        Opens a file dialog that allows users to select .xlsx files only.
    """
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx")],
        defaultextension=".xlsx",
    )
    return file_path


def update_progress(progress_bar: ttk.Progressbar, current: int, total: int) -> None:
    """Update the progress bar.

    Args:
        progress_bar: The progress bar widget to update
        current: Current sheet number (0-based)
        total: Total number of sheets

    Note:
        Calculates the percentage and updates the progress bar value.
    """
    if total > 0:
        percent = int((current / total) * 100)
        progress_bar["value"] = percent
        progress_bar.update_idletasks()


def process_file_in_background(
    file_path: str,
    progress_bar: ttk.Progressbar,
    status_label: tk.Label,
    process_button: tk.Button,
    results_text: tk.Text,
) -> None:
    """Process Excel file in a background thread.

    Args:
        file_path: Path to the Excel file
        progress_bar: Progress bar widget
        status_label: Label to show current status
        process_button: Button to re-enable when done
        results_text: Text widget to display results

    Note:
        1. Run processing in a separate thread
        2. Update progress bar and status label
        3. Show results in the text widget
        4. Handle any errors
    """

    def process_in_thread():
        try:
            process_button.config(state="disabled")
            status_label.config(text="Processing...")
            progress_bar["value"] = 0

            def progress_callback(current, total, sheet_name):
                status_label.config(text=f"Processing sheet: {sheet_name}")
                update_progress(progress_bar, current, total)

            results = process_excel_file(file_path, progress_callback)

            progress_bar["value"] = 100
            status_label.config(text="Processing complete!")

            # Show results in text widget
            results_text.delete("1.0", tk.END)
            results_text.insert(tk.END, "Row counts per sheet:\n")
            results_text.insert(tk.END, "-" * 30 + "\n")
            for sheet_name, row_count in results.items():
                results_text.insert(tk.END, f"{sheet_name}: {row_count} rows\n")

            total_rows = sum(results.values())
            results_text.insert(tk.END, "-" * 30 + "\n")
            results_text.insert(tk.END, f"Total rows: {total_rows}\n")
            results_text.insert(tk.END, f"Total sheets: {len(results)}\n")

        except Exception as e:
            status_label.config(text="Error occurred!")
            results_text.delete("1.0", tk.END)
            results_text.insert(tk.END, f"Error processing file:\n{str(e)}")

        finally:
            process_button.config(state="normal")

    threading.Thread(target=process_in_thread, daemon=True).start()


def create_main_window() -> tk.Tk:
    """Create the main application window.

    Returns:
        tk.Tk: The main window

    Note:
        create a window with title "XLSX Reader"
        and appropriate size (e.g., 500x400 to fit results display).
    """
    root = tk.Tk()
    root.title("XLSX Reader")
    root.geometry("500x400")
    root.resizable(True, True)
    return root


def run_app() -> None:
    """Run the main application.

    Note:
        1. Create the main window
        2. Add a "Select Excel File" button
        3. Add a progress bar
        4. Add a status label
        5. Add a text widget to display results
        6. Start the tkinter main loop
    """
    root = create_main_window()

    # Create and pack widgets
    # Button frame
    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    select_button = tk.Button(
        button_frame, text="Select Excel File", font=("Arial", 12), bg="lightblue", width=20
    )
    select_button.pack()

    # Progress frame
    progress_frame = tk.Frame(root)
    progress_frame.pack(pady=5)

    progress_bar = ttk.Progressbar(progress_frame, length=400, mode="determinate")
    progress_bar.pack()

    # Status frame
    status_frame = tk.Frame(root)
    status_frame.pack(pady=5)

    status_label = tk.Label(status_frame, text="Select an Excel file to begin", font=("Arial", 10))
    status_label.pack()

    # Results frame
    results_frame = tk.Frame(root)
    results_frame.pack(pady=10, padx=20, fill="both", expand=True)

    results_label = tk.Label(results_frame, text="Results:", font=("Arial", 10, "bold"))
    results_label.pack(anchor="w")

    # Text widget with scrollbar
    text_frame = tk.Frame(results_frame)
    text_frame.pack(fill="both", expand=True)

    results_text = tk.Text(text_frame, height=12, width=50, font=("Courier", 10), wrap=tk.WORD)

    scrollbar = tk.Scrollbar(text_frame, orient="vertical", command=results_text.yview)
    results_text.configure(yscrollcommand=scrollbar.set)

    results_text.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Configure button click event
    def on_select_file():
        file_path = select_excel_file()
        if file_path:
            process_file_in_background(
                file_path, progress_bar, status_label, select_button, results_text
            )

    select_button.config(command=on_select_file)

    # Start the application
    root.mainloop()
