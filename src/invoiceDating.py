from __future__ import annotations

import os
import random
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import List

import win32com.client as win32  # type: ignore


@dataclass(frozen=True)
class InvoiceDatesConfig:
    # Page anchors (page 1 and page 2) for invoice date cell
    first_page_invoice_cell: str = "K12"
    second_page_invoice_cell: str = "K61"

    # Expiration is always 1 row below invoice date in your template
    expiration_row_offset: int = 1

    # How many days after invoice date is the expiration date
    expiration_days: int = 30

    # Random day-step behavior for invoice dates
    # We step forward by 0..max_step_days days (weekday-only; repeats allowed)
    max_step_days: int = 11

    # Hard cap, like your other modules
    max_pages: int = 50


def apply_invoice_and_expiration_dates(
    file_path: str,
    *,
    total_pages: int,
    start_date: date,
    end_date: date,
    config: InvoiceDatesConfig = InvoiceDatesConfig(),
    sheet_index: int = 1,
    visible: bool = False,
) -> None:
    """
    Writes invoice date + expiration date for each page.

    Requirements enforced:
      - Input start_date and end_date must be weekdays (Mon-Fri).
      - Written dates are weekdays only (never weekends).
      - Invoice dates advance randomly but never exceed end_date.
      - Dates may repeat if necessary.
      - Excel number format: DD/MM/YYYY.

    Expiration date rule (assumption):
      expiration = invoice_date + config.expiration_days, moved forward to next weekday if weekend.
    """
    if total_pages <= 0:
        return
    if start_date.weekday() >= 5 or end_date.weekday() >= 5:
        raise ValueError("start_date and end_date must be weekdays (Mon–Fri).")
    if end_date < start_date:
        raise ValueError("end_date must be the same or after start_date.")

    pages_to_apply = min(total_pages, config.max_pages)

    abs_path = os.path.abspath(file_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"Excel file not found: {abs_path}")

    invoice_col, first_row = _cell_to_col_row(config.first_page_invoice_cell)
    _, second_row = _cell_to_col_row(config.second_page_invoice_cell)

    page_row_step = second_row - first_row
    if page_row_step <= 0:
        raise ValueError(
            f"Invalid page stride: '{config.second_page_invoice_cell}' must be below '{config.first_page_invoice_cell}'."
        )

    # Precompute all allowed weekdays in the range (for fallback sampling / repeats)
    allowed_weekdays = _weekdays_in_range(start_date, end_date)
    if not allowed_weekdays:
        raise ValueError("No weekdays available in the provided date range.")

    invoice_dates = _generate_random_weekday_dates(
        count=pages_to_apply,
        start_date=start_date,
        end_date=end_date,
        allowed_weekdays=allowed_weekdays,
        max_repeats_per_date=3,
    )

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = visible
    excel.DisplayAlerts = False

    wb = None
    try:
        wb = excel.Workbooks.Open(abs_path)
        ws = wb.Worksheets(sheet_index)

        for page_index, inv_date in enumerate(invoice_dates):
            inv_row = first_row + page_index * page_row_step
            exp_row = inv_row + config.expiration_row_offset

            exp_date = _add_days_adjust_weekday(inv_date, config.expiration_days)

            inv_cell = ws.Cells(inv_row, invoice_col)
            exp_cell = ws.Cells(exp_row, invoice_col)

            # Force DD/MM/YYYY display
            inv_cell.NumberFormat = "dd/mm/yyyy"
            exp_cell.NumberFormat = "dd/mm/yyyy"

            # COM likes datetime, not date
            inv_cell.Value = datetime(inv_date.year, inv_date.month, inv_date.day)
            exp_cell.Value = datetime(exp_date.year, exp_date.month, exp_date.day)

        wb.Save()
        print(f"Applied invoice + expiration dates to {pages_to_apply} page(s) in: {file_path}")

    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        excel.Quit()


def _generate_random_weekday_dates(
    *,
    count: int,
    start_date: date,
    end_date: date,
    allowed_weekdays: List[date],
    max_repeats_per_date: int = 3,
) -> List[date]:
    """
    Generate a monotonic (non-decreasing) sequence of invoice dates with controlled repeats.

    Constraints:
      - Dates never go backwards.
      - Repeats are contiguous blocks (by construction).
      - No date repeats more than `max_repeats_per_date` times.
      - Repeats are distributed across the range (not just the last date).

    Note:
      - If count > max_repeats_per_date * number_of_weekdays_in_range, it's impossible.
        In that case we raise a clear ValueError.

    Args:
        count: number of pages to produce dates for.
        start_date, end_date: inclusive range boundaries (weekdays).
        allowed_weekdays: sorted list of weekdays in the range (from _weekdays_in_range()).
        max_repeats_per_date: maximum times a single date can appear (default 3).

    Returns:
        List[date] of length `count`.
    """
    if count <= 0:
        return []

    # Ensure allowed_weekdays covers start..end properly
    if not allowed_weekdays:
        raise ValueError("No weekdays available in the provided date range.")

    # Find first index >= start_date, and last index <= end_date
    start_idx = None
    end_idx = None
    for i, d in enumerate(allowed_weekdays):
        if start_idx is None and d >= start_date:
            start_idx = i
        if d <= end_date:
            end_idx = i

    if start_idx is None or end_idx is None or start_idx > end_idx:
        raise ValueError("Date range does not contain any valid weekdays.")

    dates = allowed_weekdays[start_idx : end_idx + 1]
    m = len(dates)

    max_possible = max_repeats_per_date * m
    if count > max_possible:
        raise ValueError(
            f"Cannot assign {count} pages with max {max_repeats_per_date} repeats per date "
            f"over only {m} weekday(s). Maximum possible is {max_possible}."
        )

    # If we have fewer pages than available dates, we can choose an increasing subset.
    # This keeps monotonic behavior and avoids weird repetition.
    if count <= m:
        # Pick `count` indices increasing, biased toward early/middle (more realistic),
        # but always monotonic and unique.
        chosen = sorted(random.sample(range(m), count))
        return [dates[i] for i in chosen]

    # Otherwise, we must repeat some dates. Start with 1 occurrence of each date.
    repeats = [1] * m
    extras = count - m  # how many additional occurrences to distribute

    # Weighted distribution (bell curve) so repeats spread across the range.
    # Not a strict normal pdf, but a smooth peak around the middle.
    weights = _bell_weights(m)

    # Allocate extras while respecting max repeats per date.
    # This guarantees: no date repeats more than max_repeats_per_date.
    while extras > 0:
        idx = _weighted_choice_index(weights, repeats, max_repeats_per_date)
        repeats[idx] += 1
        extras -= 1

    # Expand into monotonic list with contiguous repeats per date.
    out: List[date] = []
    for d, r in zip(dates, repeats):
        out.extend([d] * r)

    # Should match exactly, but trim defensively.
    return out[:count]


def _bell_weights(m: int) -> List[float]:
    """
    Create bell-shaped weights across positions 0..m-1.
    This encourages repeats to spread around the middle, not just the end.
    """
    if m <= 1:
        return [1.0]

    mid = (m - 1) / 2.0
    sigma = max(1.0, m / 4.0)  # wider bell for longer ranges

    weights = []
    for i in range(m):
        x = (i - mid) / sigma
        w = pow(2.718281828, -0.5 * x * x)  # exp(-0.5*x^2) without importing math
        weights.append(w)

    # Avoid zero weights
    min_w = min(weights)
    if min_w <= 0:
        weights = [w - min_w + 1e-6 for w in weights]
    return weights


def _weighted_choice_index(
    weights: List[float],
    repeats: List[int],
    max_repeats_per_date: int,
) -> int:
    """
    Choose an index using weights, but only among indices that haven't hit max repeats.
    """
    eligible_indices = [i for i, r in enumerate(repeats) if r < max_repeats_per_date]
    if not eligible_indices:
        # Shouldn't happen if count <= max_possible, but safe guard.
        raise ValueError("No eligible dates left to repeat (hit max repeats everywhere).")

    eligible_weights = [weights[i] for i in eligible_indices]
    chosen = random.choices(eligible_indices, weights=eligible_weights, k=1)[0]
    return chosen


def _weekdays_in_range(start: date, end: date) -> List[date]:
    out: List[date] = []
    cur = start
    while cur <= end:
        if cur.weekday() < 5:
            out.append(cur)
        cur += timedelta(days=1)
    return out


def _add_days_adjust_weekday(d: date, days: int) -> date:
    return _adjust_forward_to_weekday(d + timedelta(days=days))


def _adjust_forward_to_weekday(d: date) -> date:
    # 5=Sat, 6=Sun
    while d.weekday() >= 5:
        d += timedelta(days=1)
    return d


def _cell_to_col_row(cell: str) -> tuple[int, int]:
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
