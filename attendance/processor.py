"""
attendance/processor.py
=======================
Converts a flat list of raw Hikvision event dicts into a
pivot-style pandas DataFrame ready for Excel export.

Responsibilities
----------------
- Parse raw event records into a tidy DataFrame.
- Group by (date, employee) and compute earliest IN / latest OUT.
- Calculate shift hours using ``attendance.calculator``.
- Pivot the tidy frame into a multi-column layout (one employee
  = three columns: IN, Out, Hours).
- Append a TOTAL row at the bottom of the sheet.
"""

from typing import Any

import pandas as pd

from attendance.calculator import calc_shift_minutes, minutes_to_hm, format_time


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def build_attendance_dataframe(
    raw_events: list[dict[str, Any]],
) -> tuple[pd.DataFrame, list[str]]:
    """
    Transform raw API event records into a pivot DataFrame.

    Parameters
    ----------
    raw_events:
        Filtered list returned by ``hikvision.fetch_events``.

    Returns
    -------
    (final_df, unique_employees)
        ``final_df``  – multi-column DataFrame indexed by date string,
                        with a TOTAL row appended at the bottom.
        ``unique_employees`` – sorted list of employee column labels,
                              needed by the Excel exporter for header
                              rendering.
    """
    df_raw = _parse_raw(raw_events)
    tidy = _build_tidy(df_raw)
    final_df, employees = _pivot(tidy)
    final_df = _append_totals(final_df, tidy, employees)
    return final_df, employees


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

def _parse_raw(raw_events: list[dict]) -> pd.DataFrame:
    """Parse the flat event list and add computed columns."""
    df = pd.DataFrame(raw_events)

    df["datetime"] = pd.to_datetime(df["time"])
    df["Date"] = df["datetime"].dt.strftime("%Y-%m-%d")

    # Employee label: "Name EmployeeID"
    df["Employee_Col"] = (
        df["name"].fillna("Unknown") + " " + df["employeeNoString"]
    )

    return df


def _build_tidy(df: pd.DataFrame) -> pd.DataFrame:
    """
    Collapse each (date, employee) group to a single row containing
    earliest IN, latest OUT, shift hours, and total minutes.
    """
    rows = []

    for (date_val, employee), group in df.groupby(["Date", "Employee_Col"]):
        check_ins = [
            row["datetime"]
            for _, row in group.iterrows()
            if str(row.get("attendanceStatus", "")).lower() == "checkin"
        ]
        check_outs = [
            row["datetime"]
            for _, row in group.iterrows()
            if str(row.get("attendanceStatus", "")).lower() == "checkout"
        ]

        in_time_str = format_time(min(check_ins)) if check_ins else ""
        out_time_str = format_time(max(check_outs)) if check_outs else ""

        total_minutes = 0
        hours_str = ""

        if check_ins and check_outs:
            total_minutes = calc_shift_minutes(min(check_ins), max(check_outs))
            hours_str = minutes_to_hm(total_minutes)

        rows.append({
            "Date": date_val,
            "Employee_Col": employee,
            "IN": in_time_str,
            "Out": out_time_str,
            "Hours": hours_str,
            "TotalMinutes": total_minutes,
        })

    return pd.DataFrame(rows)


def _pivot(tidy: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """
    Build a wide-format DataFrame: one date per row, three columns
    (IN / Out / Hours) per employee.
    """
    unique_dates = sorted(tidy["Date"].unique())
    unique_employees = sorted(tidy["Employee_Col"].unique())

    columns_tuples = []
    for emp in unique_employees:
        columns_tuples += [(emp, "IN"), (emp, "Out"), (emp, "Hours")]

    multi_cols = pd.MultiIndex.from_tuples(columns_tuples)
    final_df = pd.DataFrame(index=unique_dates, columns=multi_cols)

    for _, row in tidy.iterrows():
        emp = row["Employee_Col"]
        date = row["Date"]
        final_df.loc[date, (emp, "IN")] = row["IN"]
        final_df.loc[date, (emp, "Out")] = row["Out"]
        final_df.loc[date, (emp, "Hours")] = row["Hours"]

    return final_df.fillna(""), unique_employees


def _append_totals(
    final_df: pd.DataFrame,
    tidy: pd.DataFrame,
    employees: list[str],
) -> pd.DataFrame:
    """Compute per-employee monthly totals and append a TOTAL row."""
    totals_row: dict = {}

    for emp in employees:
        emp_rows = tidy[tidy["Employee_Col"] == emp]
        total_minutes = int(emp_rows["TotalMinutes"].sum())

        totals_row[(emp, "IN")] = ""
        totals_row[(emp, "Out")] = "TOTAL"
        totals_row[(emp, "Hours")] = minutes_to_hm(total_minutes)

    final_df.loc["TOTAL"] = totals_row
    return final_df
