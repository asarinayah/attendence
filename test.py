#!/usr/bin/env python3
"""
Convert raw 'Attendance' export text into a clean Excel workbook, with useful tweaks.

Parses lines like:
    <Attendance>: 13 : 2025-09-01 07:18:32 (1,  0)

Default output columns:
    employee_id | timestamp | date | time | signal_1 | signal_2

Options:
    --dedupe-per-minute      Keep only the earliest scan per employee per minute.
"""

import re
import sys
import pathlib
import argparse
from typing import List, Tuple

import pandas as pd


PATTERN = re.compile(
    r"[<]*Attendance[>]*:\s*(\d+)\s*:\s*"
    r"([0-9]{4}-[0-9]{2}-[0-9]{2}\s+[0-9]{2}:[0-9]{2}:[0-9]{2})\s*"
    r"\(\s*([0-9]+)\s*,\s*([0-9]+)\s*\)\.?\d*",
    flags=re.MULTILINE,
)


def parse_attendance_text(text: str) -> pd.DataFrame:
    """Parse the entire text blob into a DataFrame."""
    matches: List[Tuple[str, str, str, str]] = PATTERN.findall(text)

    if not matches:
        raise ValueError(
            "No attendance records matched. "
            "Check the input format or adjust the regex pattern."
        )

    df = pd.DataFrame(matches, columns=["employee_id", "timestamp", "signal_1", "signal_2"])

    # Types
    df["employee_id"] = df["employee_id"].astype(int)
    df["timestamp"] = pd.to_datetime(df["timestamp"], errors="coerce")
    df["signal_1"] = pd.to_numeric(df["signal_1"], errors="coerce").astype("Int64")
    df["signal_2"] = pd.to_numeric(df["signal_2"], errors="coerce").astype("Int64")

    # Drop rows with bad timestamps, sort
    df = df.dropna(subset=["timestamp"])
    df = df.sort_values(["timestamp", "employee_id"]).reset_index(drop=True)

    return df


def split_datetime(df: pd.DataFrame) -> pd.DataFrame:
    """Add 'date' and 'time' columns derived from 'timestamp'."""
    df = df.copy()
    df["date"] = df["timestamp"].dt.date
    df["time"] = df["timestamp"].dt.time
    # Reorder for readability
    cols = ["employee_id", "timestamp", "date", "time", "signal_1", "signal_2"]
    return df[cols]


def dedupe_per_minute(df: pd.DataFrame) -> pd.DataFrame:
    """
    Keep the earliest scan per employee per minute.

    Implementation: floor timestamp to minute and drop duplicates keeping first.
    """
    df = df.copy()
    df["minute"] = df["timestamp"].dt.floor("T")
    df = (
        df.sort_values(["employee_id", "timestamp"])
          .drop_duplicates(subset=["employee_id", "minute"], keep="first")
          .drop(columns=["minute"])
          .sort_values(["timestamp", "employee_id"])
          .reset_index(drop=True)
    )
    return df


def convert_raw_to_excel(input_path: pathlib.Path, output_path: pathlib.Path, use_dedupe_minute: bool) -> None:
    """Read raw file, parse, tweak, and write Excel workbook."""
    text = input_path.read_text(encoding="utf-8", errors="replace")
    df = parse_attendance_text(text)

    # Always split timestamp into date/time for readability
    if use_dedupe_minute:
        df = dedupe_per_minute(df)

    df = split_datetime(df)

    with pd.ExcelWriter(output_path, engine="xlsxwriter",
                        datetime_format="yyyy-mm-dd hh:mm:ss") as writer:
        df.to_excel(writer, index=False, sheet_name="Attendance")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Convert raw attendance export into a clean .xlsx")
    p.add_argument("input_raw_file", type=pathlib.Path, help="Path to raw attendance file")
    p.add_argument("output_xlsx", type=pathlib.Path, help="Output .xlsx path")
    p.add_argument("--dedupe-per-minute", action="store_true",
                   help="Keep only the earliest scan per employee per minute")
    return p


def main():
    parser = build_parser()
    args = parser.parse_args()

    in_path: pathlib.Path = args.input_raw_file.expanduser().resolve()
    out_path: pathlib.Path = args.output_xlsx.expanduser().resolve()

    if not in_path.exists() or not in_path.is_file():
        print(f"Input file not found: {in_path}")
        sys.exit(2)

    out_path.parent.mkdir(parents=True, exist_ok=True)

    convert_raw_to_excel(in_path, out_path, use_dedupe_minute=args.dedupe_per_minute)
    print(f"Done. Wrote: {out_path}")


if __name__ == "__main__":
    main()
