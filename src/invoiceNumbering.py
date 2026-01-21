from __future__ import annotations

import os
from typing import Optional

import win32com.client as win32  # type: ignore

def _cell_to_col_row(cell: str) -> tuple[int, int]:
    """
    Convert an Excel A1 cell reference (e.g. 'E10') into (col_number, row_number).
    Column is 1-based (A=1), row is int.
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

def apply_invoice_numbers(
    file_path: str,
    *,
    start_number: int,
    total_pages: int,
    first_page_cell: str = "E10",
    second_page_cell: str = "E59",
    sheet_index: int = 1,
    visible: bool = False,
    max_pages: int = 50,   # <-- NEW
) -> None:
    """
    Writes sequential invoice numbers into a fixed template on one worksheet.

    Applies numbering to at most `max_pages` pages.
    """
    if start_number <= 0:
        raise ValueError("start_number must be a positive integer.")

    if total_pages <= 0:
        return

    pages_to_apply = min(total_pages, max_pages)

    abs_path = os.path.abspath(file_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"Excel file not found: {abs_path}")

    invoice_col, first_row = _cell_to_col_row(first_page_cell) # type: ignore
    _, second_row = _cell_to_col_row(second_page_cell) # type: ignore

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

        for page_index in range(pages_to_apply):
            invoice_number = start_number + page_index
            row = first_row + page_index * page_row_step
            ws.Cells(row, invoice_col).Value = invoice_number

        wb.Save()
        print(
            f"Applied invoice numbers {start_number}..{start_number + pages_to_apply - 1} "
            f"({pages_to_apply} page(s)) to {file_path}"
        )

    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        excel.Quit()
