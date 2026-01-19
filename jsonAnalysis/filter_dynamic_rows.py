#!/usr/bin/env python3
"""
Take CSV file(s), keep only rows whose object_name is 'dynamic',
and write a single Excel file with filtered data.

INPUT:
- A single CSV file OR
- A directory containing multiple CSV files

OUTPUT:
- One Excel file
- One sheet per input CSV
"""

import argparse
from pathlib import Path
import pandas as pd


# <<< EDIT THIS if you want other columns >>>
COLUMNS_TO_KEEP = [
    "slide_id",
    "worker_id",
    "assignment_id",
    "object_name",
    "moves_count",
    "end_x",
    "end_y",
    "optimal_2d",
    "real_2d",
    "ratio_2d",
]


def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument(
        "input_path",
        help="Path to a CSV file OR a directory containing CSV files",
    )
    p.add_argument(
        "-o",
        "--output-file",
        default="motion_intent_analysis.xlsx",
        help="Output Excel file (default: motion_intent_analysis.xlsx)",
    )
    p.add_argument(
        "--filter-column",
        default="object_name",
        help="Column that contains the word 'dynamic' (default: object_name)",
    )
    return p.parse_args()


def main():
    args = parse_args()
    input_path = Path(args.input_path)

    if not input_path.exists():
        raise SystemExit(f"Input path not found: {input_path}")

    # Collect CSV files
    if input_path.is_file():
        if input_path.suffix.lower() != ".csv":
            raise SystemExit("Input file must be a .csv")
        csv_files = [input_path]
    else:
        csv_files = sorted(input_path.glob("*.csv"))

    if not csv_files:
        raise SystemExit("No CSV files found")

    with pd.ExcelWriter(args.output_file, engine="openpyxl") as writer:
        for csv_path in csv_files:
            df = pd.read_csv(csv_path)

            if args.filter_column not in df.columns:
                raise SystemExit(
                    f"Column '{args.filter_column}' not found in {csv_path.name}.\n"
                    f"Available columns: {list(df.columns)}"
                )

            # keep only rows where the filter column is 'dynamic'
            mask = df[args.filter_column].astype(str).str.lower().eq("dynamic")
            filtered = df.loc[mask]

            # keep only desired columns
            cols = [c for c in COLUMNS_TO_KEEP if c in filtered.columns]
            filtered = filtered[cols]

            # sheet name = CSV file name (Excel limit: 31 chars)
            sheet_name = csv_path.stem[:31] or "data"
            filtered.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"Done. Wrote filtered rows to: {args.output_file}")


if __name__ == "__main__":
    main()

#############
# HOW TO RUN:
#
# 1) Run on a SINGLE CSV file:
#
#   python3 filter_dynamic_rows.py \
#     /full/path/to/slide01.csv \
#     --output-file motion_intent_analysis.xlsx
#
# 2) Run on a DIRECTORY of CSV files:
#
#   python3 filter_dynamic_rows.py \
#     /full/path/to/distances_csv \
#     --output-file motion_intent_analysis.xlsx
#
# ------------------------------------------------------------
#
# OPTIONAL ARGUMENTS:
#
#   --output-file
#       Path/name of the Excel file to create.
#       Default: motion_intent_analysis.xlsx
#
#   --filter-column
#       Name of the column used to filter rows.
#       Only rows where this column equals "dynamic" are kept.
#       Default: object_name
#
#############
