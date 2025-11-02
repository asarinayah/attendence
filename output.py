#!/usr/bin/env python3
"""
attendance_to_sample_format.py

Create a monthly attendance workbook (Sample-like format):
- First two columns: "Employee Code", "Employee"
- Date columns for every calendar day in each month found in the attendance file
- Cells contain "P" if employee has attendance for that date, else blank
- One sheet per month (e.g., "Sep-2025")

Inputs:
  1) --attendance PATH   : cleaned_attendance.xlsx (or similar)
     - Must contain 'employee_id' and either 'date' or 'timestamp' column.
  2) --names PATH        : Names.xlsx (first column is employee code, second is employee name)
  3) --output PATH       : Output .xlsx file path (default: Attendance_By_Month.xlsx)
  4) --month YYYY-MM     : (Optional) Restrict to a single month, e.g. 2025-09
  5) --engine NAME       : (Optional) pandas ExcelWriter engine (default: xlsxwriter)

Examples:
  python attendance_to_sample_format.py --attendance cleaned_attendance.xlsx --names Names.xlsx --output Attendance_By_Month.xlsx
  python attendance_to_sample_format.py --attendance cleaned_attendance.xlsx --names Names.xlsx --output Sep2025.xlsx --month 2025-09

Dependencies:
  pip install pandas openpyxl xlsxwriter
"""

import argparse
from calendar import monthrange
from datetime import date, datetime
from typing import List, Tuple, Optional, Set
import pandas as pd


def parse_args():
    p = argparse.ArgumentParser(description="Convert attendance logs to Sample-like monthly sheets.")
    p.add_argument("--attendance", required=True, help="Path to attendance Excel file (must have employee_id and date/timestamp).")
    p.add_argument("--names", required=True, help="Path to Names Excel file (first column = Employee Code, second = Employee Name).")
    p.add_argument("--output", default="Attendance_By_Month.xlsx", help="Path to output Excel workbook.")
    p.add_argument("--month", default=None, help="Optional month filter in YYYY-MM format (e.g., 2025-09).")
    p.add_argument("--engine", default="xlsxwriter", help="Pandas ExcelWriter engine (xlsxwriter or openpyxl).")
    return p.parse_args()


def normalize_attendance(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip().lower() for c in df.columns]
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.date
    elif "timestamp" in df.columns:
        df["date"] = pd.to_datetime(df["timestamp"], errors="coerce").dt.date
    else:
        raise ValueError("Attendance file must contain 'date' or 'timestamp' column.")
    if "employee_id" not in df.columns:
        candidates = [c for c in df.columns if "emp" in c and "id" in c]
        if not candidates:
            raise ValueError("Attendance file must contain 'employee_id' column.")
        df.rename(columns={candidates[0]: "employee_id"}, inplace=True)
    df["employee_id"] = df["employee_id"].astype(str).str.strip()
    return df


def normalize_names(df: pd.DataFrame) -> pd.DataFrame:
    if df.shape[1] < 2:
        raise ValueError("Names file must have at least two columns: [Employee Code, Employee Name].")
    cols = list(df.columns)
    df = df.rename(columns={cols[0]: "Employee Code", cols[1]: "Employee"})[["Employee Code", "Employee"]].copy()
    df["Employee Code"] = df["Employee Code"].astype(str).str.strip()
    df["Employee"] = df["Employee"].astype(str).str.strip()
    return df.drop_duplicates()


def month_key(d: date) -> Tuple[int, int]:
    return (d.year, d.month)


def iter_calendar_month(year: int, month: int) -> List[pd.Timestamp]:
    days_in_month = monthrange(year, month)[1]
    return [pd.Timestamp(date(year, month, d)) for d in range(1, days_in_month + 1)]


def build_present_set(att: pd.DataFrame, month_filter: Optional[str]) -> Tuple[Set[Tuple[str, str]], List[Tuple[int, int]]]:
    df = att.dropna(subset=["date"]).copy()
    if month_filter:
        y, m = map(int, month_filter.split("-"))
        df = df[(df["date"].map(lambda d: d.year) == y) & (df["date"].map(lambda d: d.month) == m)]
    months = sorted({month_key(d) for d in df["date"]})
    present = df.groupby(["employee_id", "date"], as_index=False).size()
    present["employee_id"] = present["employee_id"].astype(str).str.strip()
    present_set = set(zip(present["employee_id"], present["date"].astype(str)))
    return present_set, months


def make_sheet(writer: pd.ExcelWriter, roster: pd.DataFrame, year: int, month: int, present_set: Set[Tuple[str, str]]) -> None:
    all_dates = iter_calendar_month(year, month)
    wide = roster.copy()
    for d in all_dates:
        wide[d] = ""
    for idx, row in wide.iterrows():
        code = str(row["Employee Code"]).strip()
        for d in all_dates:
            if (code, str(d.date())) in present_set:
                wide.at[idx, d] = "P"
    sheet_name = pd.Timestamp(date(year, month, 1)).strftime("%b-%Y")
    wide.to_excel(writer, sheet_name=sheet_name, index=False)


def main():
    args = parse_args()
    att = pd.read_excel(args.attendance, sheet_name=0)
    names = pd.read_excel(args.names, sheet_name=0)
    att = normalize_attendance(att)
    roster = normalize_names(names)
    present_set, months = build_present_set(att, args.month)

    if not months:
        with pd.ExcelWriter(args.output, engine=args.engine, datetime_format="yyyy-mm-dd") as writer:
            roster.to_excel(writer, sheet_name="NoData", index=False)
        print(f"✅ No data found. Wrote roster-only workbook: {args.output}")
        return

    with pd.ExcelWriter(args.output, engine=args.engine, datetime_format="yyyy-mm-dd") as writer:
        for (yr, mo) in months:
            make_sheet(writer, roster, yr, mo, present_set)
    print(f"✅ Done. Created {len(months)} sheet(s): {args.output}")


if __name__ == "__main__":
    main()
