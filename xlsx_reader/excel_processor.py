"""Excel processing module for reading XLSX files.

This module provides simple functions to read Excel files and get basic
sheet information using pandas.
"""

from collections.abc import Callable
from openpyxl import load_workbook


def get_sheet_names(file_path: str) -> list[str]:
    wb = load_workbook(file_path, read_only=True)
    sheetnames = wb.sheetnames
    wb.close()  # <-- close workbook to release file handle
    return sheetnames


def get_sheet_row_count(file_path: str, sheet_name: str) -> int:
    wb = load_workbook(file_path, read_only=True)
    sheet = wb[sheet_name]
    row_count = 0
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if any(cell is not None and str(cell).strip() != "" for cell in row):
            row_count += 1
    wb.close()  # <-- close workbook here
    return row_count


def process_excel_file(
    file_path: str, progress_callback: Callable[[int, int, str], None] | None = None
) -> dict[str, int]:
    """Process an Excel file and return row counts for each sheet."""
    sheet_names = get_sheet_names(file_path)
    total_sheets = len(sheet_names)
    results = {}

    for current_index, sheet_name in enumerate(sheet_names):
        if progress_callback:
            progress_callback(current_index, total_sheets, sheet_name)

        row_count = get_sheet_row_count(file_path, sheet_name)
        results[sheet_name] = row_count

    return results
