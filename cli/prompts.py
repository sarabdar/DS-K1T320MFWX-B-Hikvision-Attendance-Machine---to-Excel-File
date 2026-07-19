"""
cli/prompts.py
==============
All interactive CLI prompts that the user sees at startup.

Returns plain data (no side effects beyond printing) so that
``main.py`` can pass the result straight to business-logic modules.
"""

import calendar
import sys
from dataclasses import dataclass
from datetime import datetime


# ---------------------------------------------------------------------------
# Data class returned to the caller
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class MonthSelection:
    """Holds the user's chosen month and year together with derived metadata."""

    year: int
    month: int

    @property
    def month_name(self) -> str:
        return calendar.month_name[self.month]

    @property
    def last_day(self) -> int:
        return calendar.monthrange(self.year, self.month)[1]

    @property
    def start_time(self) -> str:
        return f"{self.year}-{self.month:02d}-01T00:00:00+05:00"

    @property
    def end_time(self) -> str:
        return f"{self.year}-{self.month:02d}-{self.last_day:02d}T23:59:59+05:00"

    @property
    def output_filename(self) -> str:
        return f"Hikvision_Attendance_{self.month_name}_{self.year}.xlsx"


# ---------------------------------------------------------------------------
# Month name → number mapping
# ---------------------------------------------------------------------------

_MONTH_MAP: dict[str, int] = {
    "january": 1, "february": 2, "march": 3,
    "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "september": 9,
    "october": 10, "november": 11, "december": 12,
}


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def prompt_month_selection() -> MonthSelection:
    """
    Interactively ask the user which month/year to generate.

    Defaults to the current month when the user confirms.
    Exits with an informative message if an invalid month name is given.

    Returns
    -------
    MonthSelection
        Immutable object with all derived date strings pre-computed.
    """
    now = datetime.now()
    current_year = now.year
    current_month = now.month
    current_month_name = calendar.month_name[current_month]

    print(
        f"\nYou want to make Excel attendance sheet "
        f"for {current_month_name} {current_year}?"
    )

    choice = input("Type y for yes and n for no: ").strip().lower()

    if choice == "y":
        return MonthSelection(year=current_year, month=current_month)

    # ── Custom month ──────────────────────────────────────────────────────
    print("\nOkay, which month?")
    print("Example: april")

    user_month = input("Month name: ").strip().lower()

    if user_month not in _MONTH_MAP:
        print(f"\nInvalid month entered: '{user_month}'")
        sys.exit(1)

    selected_month = _MONTH_MAP[user_month]

    year_input = input(
        f"Enter year (press Enter for {current_year}): "
    ).strip()

    selected_year = int(year_input) if year_input else current_year

    return MonthSelection(year=selected_year, month=selected_month)


def print_job_summary(selection: MonthSelection) -> None:
    """Print a formatted summary of the upcoming export job."""
    print("\n==========================================")
    print(f"Generating sheet for: {selection.month_name} {selection.year}")
    print(f"Start Time: {selection.start_time}")
    print(f"End Time:   {selection.end_time}")
    print(f"Output:     {selection.output_filename}")
    print("==========================================\n")
