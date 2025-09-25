"""Excel processing module for reading XLSX files."""

from collections.abc import Callable

from openpyxl import load_workbook


def get_sheet_names(file_path: str) -> list[str]:
    """Get all sheet names from an Excel file."""
    wb = load_workbook(file_path, read_only=True)
    sheet_names = wb.sheetnames
    wb.close()  # Important: close the workbook to release file lock
    return sheet_names


def get_sheet_row_count(file_path: str, sheet_name: str) -> int:
    """Get the number of rows in a specific sheet (excluding header)."""
    wb = load_workbook(file_path, read_only=True)
    ws = wb[sheet_name]

    # Count rows that contain data (ignoring completely empty rows)
    row_count = 0
    for row in ws.iter_rows(values_only=True):
        if any(cell is not None for cell in row):  # row has data
            row_count += 1

    wb.close()  # Close workbook to release file lock

    # Exclude header row if there is at least one row
    return max(row_count - 1, 0)


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
