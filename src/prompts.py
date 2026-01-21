from __future__ import annotations
from dataclasses import dataclass
from datetime import date

def _prompt_invoice_start() -> int:
    """
    Prompt user for the first invoice number.

    Returns:
        Positive integer invoice number.
    """
    while True:
        raw = input("Start invoice number (e.g., 804): ").strip()
        try:
            value = int(raw)
            if value <= 0:
                raise ValueError
            return value
        except ValueError:
            print("Please enter a positive integer (e.g., 804).")


def _prompt_control_start() -> int:
    """
    Prompt user for the first control number.

    Control numbers are displayed as 6 digits (e.g., 000968), but the user can enter
    them with or without leading zeros.

    Returns:
        Integer in range [0, 999999].
    """
    while True:
        raw = input("Start control number (6 digits, e.g., 000968): ").strip()

        # Allow inputs like "000968" or "968"
        if not raw.isdigit():
            print("Please enter digits only (e.g., 000968).")
            continue

        value = int(raw)
        if 0 <= value <= 999_999:
            return value

        print("Control number must be between 000000 and 999999.")


@dataclass(frozen=True)
class DateRange:
    start: date
    end: date


def prompt_invoice_date_range() -> DateRange:
    """
    Ask user for a start/end date in DD/MM/YYYY.
    Both must be weekdays (Mon-Fri). End must be >= start.
    """
    while True:
        start = _prompt_weekday_date("Start date (DD/MM/YYYY): ")
        end = _prompt_weekday_date("End date (DD/MM/YYYY): ")

        if end < start:
            print("End date must be the same or after the start date.")
            continue

        return DateRange(start=start, end=end)


def _prompt_weekday_date(prompt: str) -> date:
    while True:
        raw = input(prompt).strip()
        parsed = _parse_ddmmyyyy(raw)
        if parsed is None:
            print("Invalid date format. Please use DD/MM/YYYY (e.g., 21/01/2026).")
            continue

        if parsed.weekday() >= 5:
            print("That date is on a weekend. Please enter a weekday (Mon–Fri).")
            continue

        return parsed


def _parse_ddmmyyyy(raw: str) -> date | None:
    parts = raw.split("/")
    if len(parts) != 3:
        return None
    try:
        d = int(parts[0])
        m = int(parts[1])
        y = int(parts[2])
        return date(year=y, month=m, day=d)
    except ValueError:
        return None
