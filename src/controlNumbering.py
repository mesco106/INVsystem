from __future__ import annotations

import os
import random
from typing import Tuple

import win32com.client as win32  # type: ignore


def apply_control_numbers(
    file_path: str,
    *,
    start_number: int,
    total_pages: int,
    first_page_cell: str = "K6",
    second_page_cell: str = "K55",
    sheet_index: int = 1,
    visible: bool = False,
    max_pages: int = 50,
    min_jump: int = 1,
    max_jump: int = 11,
) -> None:
    """
    Applies a red, 6-digit control number to each invoice page.

    - Page anchor is derived from the row difference between first_page_cell and second_page_cell.
    - The number increments randomly by a jump in [min_jump, max_jump].
    - Always displayed as 6 digits (leading zeros preserved).
    - Text is always red.

    Args:
        file_path: Existing .xlsx file.
        start_number: Control number to place on page 1 (e.g. 968 -> displays as 000968).
        total_pages: Pages that exist/should be numbered.
        first_page_cell: Control number cell for page 1 (default K6).
        second_page_cell: Control number cell for page 2 (default K55).
        sheet_index: 1-based worksheet index.
        visible: Show Excel UI (debugging).
        max_pages: Hard cap to avoid writing too far (default 50).
        min_jump: Minimum random increment (default 1).
        max_jump: Maximum random increment (default 11).
    """
    if start_number < 0:
        raise ValueError("start_number must be >= 0.")
    if not (0 <= start_number <= 999_999):
        raise ValueError("start_number must fit in 6 digits (0..999999).")
    if total_pages <= 0:
        return
    if min_jump <= 0 or max_jump < min_jump:
        raise ValueError("Invalid jump range. Ensure 1 <= min_jump <= max_jump.")

    pages_to_apply = min(total_pages, max_pages)

    abs_path = os.path.abspath(file_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"Excel file not found: {abs_path}")

    col, first_row = _cell_to_col_row(first_page_cell)
    _, second_row = _cell_to_col_row(second_page_cell)

    page_row_step = second_row - first_row
    if page_row_step <= 0:
        raise ValueError(
            f"Invalid page stride: '{second_page_cell}' must be below '{first_page_cell}'."
        )

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = visible
    excel.DisplayAlerts = False

    wb = None
    try:
        wb = excel.Workbooks.Open(abs_path)
        ws = wb.Worksheets(sheet_index)

        current = start_number

        for page_index in range(pages_to_apply):
            row = first_row + page_index * page_row_step
            cell = ws.Cells(row, col)

            # Ensure 6 digits are displayed (leading zeros).
            cell.NumberFormat = "000000"
            cell.Value = int(current)

            # Make it red (Excel/VBA vbRed == 255).
            cell.Font.Color = 255

            # Next number: random jump 1..11
            jump = random.randint(min_jump, max_jump)
            current += jump

            if current > 999_999:
                raise ValueError(
                    f"Control number exceeded 6 digits on page {page_index + 1}. "
                    f"Current={current}. Choose a smaller start_number."
                )

        wb.Save()
        last_written = start_number  # compute last written for message
        # (We don't store all numbers; message kept simple and safe.)
        print(f"Applied {pages_to_apply} red control number(s) to {file_path}")

    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        excel.Quit()


def _cell_to_col_row(cell: str) -> Tuple[int, int]:
    """
    Convert an Excel A1 cell reference (e.g. 'K6') into (col_number, row_number).
    Column is 1-based (A=1).
    """
    cell = cell.strip().upper()
    if not cell:
        raise ValueError("Cell reference cannot be empty.")

    letters = []
    digits = []
    for ch in cell:
        if "A" <= ch <= "Z":
            if digits:
                raise ValueError(f"Invalid cell reference: {cell}")
            letters.append(ch)
        elif "0" <= ch <= "9":
            digits.append(ch)
        else:
            raise ValueError(f"Invalid cell reference: {cell}")

    if not letters or not digits:
        raise ValueError(f"Invalid cell reference: {cell}")

    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - ord("A") + 1)

    row = int("".join(digits))
    if row <= 0:
        raise ValueError(f"Invalid row in cell reference: {cell}")

    return col, row
