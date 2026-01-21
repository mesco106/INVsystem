from __future__ import annotations

import os
from typing import Any, Dict, List

import win32com.client as win32  # type: ignore


def input_products_page(
    products: List[Dict[str, Any]],
    file_path: str,
    *,
    first_page_start_row: int = 19,
    second_page_start_row: int = 68,
    start_col: int = 2,
    items_per_page: int = 9,
    row_step: int = 2,
    sheet_index: int = 1,
    visible: bool = False,
) -> None:
    """
    Writes products into a fixed invoice template on ONE worksheet, across multiple pages.

    Page anchors are determined by the template:
      - Page 1 item 1 starts at first_page_start_row
      - Page 2 item 1 starts at second_page_start_row
      - All next pages follow the same stride (page_row_step)

    Within each page, each item uses:
      - descripcion     -> (item_row,     start_col)
      - cantidad        -> (item_row,     start_col + 6)
      - precio_unitario -> (item_row,     start_col + 7)
      - precio_total    -> (item_row,     start_col + 8)
      - codigo label    -> (item_row + 1, start_col) as "NUMERO DE PARTE: <codigo>"

    Args:
        products: List of product dicts.
        file_path: Path to the Excel file (existing preferred; will create if missing).
        first_page_start_row: Row where item #1 of page 1 starts (19).
        second_page_start_row: Row where item #1 of page 2 starts (68).
        start_col: Column where descripcion is written.
        items_per_page: How many items fit in one page (9).
        row_step: Row spacing between items within the same page.
        sheet_index: 1-based Excel sheet index.
        visible: Show Excel while running (debugging).
    """
    if not products:
        raise ValueError("products is empty. Provide at least one product.")

    page_row_step = second_page_start_row - first_page_start_row
    if page_row_step <= 0:
        raise ValueError(
            f"Invalid page stride: second_page_start_row ({second_page_start_row}) "
            f"must be greater than first_page_start_row ({first_page_start_row})."
        )

    abs_path = os.path.abspath(file_path)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = visible
    excel.DisplayAlerts = False

    wb = None
    try:
        if os.path.exists(abs_path):
            wb = excel.Workbooks.Open(abs_path)
        else:
            wb = excel.Workbooks.Add()
            wb.SaveAs(abs_path)

        ws = wb.Worksheets(sheet_index)

        for idx, product in enumerate(products):
            page_index = idx // items_per_page          # 0-based page number
            slot_index = idx % items_per_page           # 0..8 within the page

            page_start_row = first_page_start_row + page_index * page_row_step
            item_row = page_start_row + slot_index * row_step

            COL_DESC = 0
            COL_QTY = 6
            COL_UNIT = 7
            COL_TOTAL = 9  # <-- K relative to start_col=2 (B)

            ws.Cells(item_row, start_col + COL_DESC).Value = product.get("descripcion", "")
            ws.Cells(item_row, start_col + COL_QTY).Value = product.get("cantidad", "")
            ws.Cells(item_row, start_col + COL_UNIT).Value = product.get("precio_unitario", "")
            ws.Cells(item_row, start_col + COL_TOTAL).Value = product.get("precio_total", "")

            codigo = product.get("codigo", "")
            ws.Cells(item_row + 1, start_col).Value = f"NUMERO DE PARTE: {codigo}"

        wb.Save()
        pages_used = (len(products) + items_per_page - 1) // items_per_page
        print(f"Wrote {len(products)} item(s) across {pages_used} page(s) into: {file_path}")

    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        excel.Quit()
