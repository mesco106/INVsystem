from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

from products import load_products
from invoicing import input_products_page
from invoiceNumbering import apply_invoice_numbers
from controlNumbering import apply_control_numbers
from invoiceDating import apply_invoice_and_expiration_dates
from prompts import _prompt_invoice_start, _prompt_control_start, prompt_invoice_date_range


# ---------------------------
# Project paths
# ---------------------------

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = PROJECT_ROOT / "data"

CSV_FILE = DATA_DIR / "processed" / "inputTable.csv"
TEMPLATE_FILE = DATA_DIR / "template" / "template.xlsx"
INVOICE_DIR = DATA_DIR / "invoice"  # <-- finished invoices go here


# ---------------------------
# Template configuration
# ---------------------------

ITEMS_PER_PAGE = 9
FIRST_PAGE_ITEMS_ROW = 19

CONTROL_FIRST_CELL = "K6"
CONTROL_SECOND_CELL = "K55"


def _compute_total_pages(num_items: int, items_per_page: int = ITEMS_PER_PAGE) -> int:
    return (num_items + items_per_page - 1) // items_per_page


def _create_invoice_from_template(
    template_path: Path,
    output_dir: Path,
    *,
    filename_prefix: str = "invoice",
) -> Path:
    """
    Copy the template file into output_dir with a unique name and return its path.
    """
    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")

    output_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = output_dir / f"{filename_prefix}_{timestamp}.xlsx"

    shutil.copy2(template_path, output_path)
    return output_path


def main() -> None:
    products = load_products(CSV_FILE)
    total_pages = _compute_total_pages(len(products))

    # Prompts (collect inputs first)
    invoice_start = _prompt_invoice_start()
    control_start = _prompt_control_start()
    date_range = prompt_invoice_date_range()

    # Create a fresh invoice file from template
    xlsx_file = _create_invoice_from_template(
        template_path=TEMPLATE_FILE,
        output_dir=INVOICE_DIR,
        filename_prefix="invoice",
    )

    # 1) Write product line items
    input_products_page(
        products=products,
        file_path=str(xlsx_file),
        first_page_start_row=FIRST_PAGE_ITEMS_ROW,
    )

    # 2) Invoice numbering
    apply_invoice_numbers(
        file_path=str(xlsx_file),
        start_number=invoice_start,
        total_pages=total_pages,
    )

    # 3) Control numbering
    apply_control_numbers(
        file_path=str(xlsx_file),
        start_number=control_start,
        total_pages=total_pages,
        first_page_cell=CONTROL_FIRST_CELL,
        second_page_cell=CONTROL_SECOND_CELL,
    )

    # 4) Dates
    apply_invoice_and_expiration_dates(
        file_path=str(xlsx_file),
        total_pages=total_pages,
        start_date=date_range.start,
        end_date=date_range.end,
    )

    print(f"\n✅ Finished invoice created: {xlsx_file}")


if __name__ == "__main__":
    main()
