"""
main.py
=======
Entry point for the Hikvision Attendance Excel Generator.


  1. CLI  →  ask the user which month to generate
  2. API  →  fetch raw attendance events from the Hikvision device
  3. ETL  →  transform events into a pivot DataFrame
  4. Export → write the styled Excel workbook to disk


  - cli/         : user interaction & prompts
  - hikvision/   : HTTP client & paginated event fetching
  - attendance/  : data transformation & hour calculations
  - export/      : openpyxl Excel styling & writing
  - config.py    : environment variables & constants
"""

import logging
import sys

from cli import prompt_month_selection, print_job_summary
from hikvision import build_session, fetch_events
from attendance import build_attendance_dataframe
from export import export_to_excel

# ── Logging ────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Pipeline
# ---------------------------------------------------------------------------

def main() -> None:
    # ── Phase 1: CLI prompt ───────────────────────────────────────────────
    selection = prompt_month_selection()
    print_job_summary(selection)

    # ── Phase 2: Fetch events from device ─────────────────────────────────
    session = build_session()

    raw_events = fetch_events(
        session=session,
        start_time=selection.start_time,
        end_time=selection.end_time,
    )

    if not raw_events:
        print("\nNo attendance logs found for the selected month.")
        sys.exit(0)

    print(f"\nProcessing {len(raw_events)} attendance logs…")

    # ── Phase 3: Transform ────────────────────────────────────────────────
    final_df, unique_employees = build_attendance_dataframe(raw_events)

    # ── Phase 4: Export ───────────────────────────────────────────────────
    export_to_excel(
        final_df=final_df,
        unique_employees=unique_employees,
        output_path=selection.output_filename,
    )

    print("\n==========================================")
    print("DONE SUCCESSFULLY!")
    print(f"Excel file created: {selection.output_filename}")
    print("==========================================")


# ---------------------------------------------------------------------------
# Guard
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    main()