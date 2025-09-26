"""Excel processing module for reading XLSX files.

This module provides simple functions to read Excel files and get basic
sheet information using pandas.
"""

from collections.abc import Callable

from openpyxl import load_workbook


def get_sheet_names(file_path: str) -> list[str]:
    """Get all sheet names from an Excel file.

    Args:
        file_path (str): Path to the Excel file (.xlsx)

    Returns:
        List[str]: List of sheet names in the workbook

    Note:
        Uses openpyxl to read the sheet names without loading all data.
    """
    workbook = load_workbook(filename=file_path, read_only=True)
    workbook.close()
    return workbook.sheetnames


def get_sheet_row_count(file_path: str, sheet_name: str) -> int:
    """Get the number of rows in a specific sheet.

    Args:
        file_path (str): Path to the Excel file (.xlsx)
        sheet_name (str): Name of the sheet to analyze

    Returns:
        int: Number of rows in the sheet

    Note:
        Uses openpyxl to count rows with data in the specified sheet.
        Excludes header row from count.
    """
    workbook = load_workbook(filename=file_path, read_only=True)
    workbook.close()
    return workbook[sheet_name].max_row - 1  # Exclude header row


def process_excel_file(
    file_path: str, progress_callback: Callable[[int, int, str], None] | None = None
) -> dict[str, int]:
    """Process an Excel file and return row counts for each sheet.

    Args:
        file_path (str): Path to the Excel file (.xlsx)
        progress_callback (Optional[Callable]): Function to call for progress updates.
            Receives (current_sheet_index, total_sheets, sheet_name)

    Returns:
        Dict[str, int]: Dictionary mapping sheet names to their row counts

    Note:
        1. Get all sheet names using get_sheet_names()
        2. For each sheet, call get_sheet_row_count()
        3. Call progress_callback if provided
        4. Return a dictionary with sheet names as keys and row counts as values
    """
    sheet_names = get_sheet_names(file_path)
    total_sheets = len(sheet_names)
    results = {}

    for current_index, sheet_name in enumerate(sheet_names):
        if progress_callback:
            progress_callback(current_index, total_sheets, sheet_name)

        row_count = get_sheet_row_count(file_path, sheet_name)
        results[sheet_name] = row_count

    return results
