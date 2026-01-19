#!/usr/bin/env python3
"""
Add per-static distance columns to an existing Excel file.

This script enriches an Excel file (per slide, per worker) by adding
distance-from-dynamic columns for each static object on the slide.

For each sheet in the input Excel:
- Assumes rows include dynamic USER-END coordinates (end_x, end_y).
- Loads the corresponding per-slide CSV file (<slide_id>.csv).
- Identifies all STATIC objects on that slide.
- For each static object S, adds a column:
    dist_to__S = Euclidean distance from dynamic USER-END to static S.


INPUT:
--input-xlsx : Excel file with dynamic rows
                 (may already include line-distance metrics).
--csv-dir    : Directory containing per-slide CSV files,
                 named exactly <slide_id>.csv.

OUTPUT:
- Excel file (default: Preferences_with_static_distances.xlsx)
- Same sheets as input, with additional static-distance columns.

Notes:
- Static object names are sanitized to be Excel-safe.
- One column is added per static object on the slide.
"""

import re
import math
import argparse
from pathlib import Path

import pandas as pd


def euclid(x1, y1, x2, y2):
    return math.hypot(x2 - x1, y2 - y1)


def sanitize_col(name: str) -> str:
    """
    Make a safe Excel column suffix from an object name.
    Keeps letters/numbers/underscore; replaces others with underscore.
    """
    s = str(name).strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^0-9A-Za-z_]", "_", s)
    s = re.sub(r"_+", "_", s)
    return s[:80]  # keep it reasonable


def load_static_csv(csv_dir: Path, slide_id: str) -> pd.DataFrame:
    """
    Expects a file named '<slide_id>.csv' inside csv_dir.
    """
    csv_path = csv_dir / f"{slide_id}.csv"
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV not found for slide_id='{slide_id}': {csv_path}")

    df = pd.read_csv(csv_path)

    needed = {"object_name", "end_x", "end_y"}
    missing = needed - set(df.columns)
    if missing:
        raise ValueError(f"{csv_path} missing columns: {sorted(missing)}")

    # keep only static objects
    df = df[df["object_name"] != "dynamic"].copy()

    # drop duplicates by name if any (keep first)
    df = df.drop_duplicates(subset=["object_name"], keep="first")

    return df


def add_static_distance_columns(sheet_df: pd.DataFrame, static_df: pd.DataFrame) -> pd.DataFrame:
    """
    For each static object, add a column dist_to__<object_name>
    computed from dynamic USER-END (end_x,end_y) to static end_x,end_y
    """
    # We only compute for rows where object_name == 'dynamic'
    if "object_name" not in sheet_df.columns:
        raise ValueError("Sheet is missing 'object_name' column")
    if "end_x" not in sheet_df.columns or "end_y" not in sheet_df.columns:
        raise ValueError("Sheet is missing 'end_x'/'end_y' columns (USER-END)")

    # Prepare output df (keep everything)
    out = sheet_df.copy()

    # Build mapping: static_name -> (x,y)
    statics = []
    for _, r in static_df.iterrows():
        statics.append((r["object_name"], float(r["end_x"]), float(r["end_y"])))

    # Create columns (vectorized per static)
    dyn_mask = out["object_name"] == "dynamic"
    dyn_end_x = out.loc[dyn_mask, "end_x"].astype(float)
    dyn_end_y = out.loc[dyn_mask, "end_y"].astype(float)

    for (name, sx, sy) in statics:
        col = f"dist_to__{sanitize_col(name)}"
        # default NaN for non-dynamic rows
        out[col] = float("nan")
        # fill only dynamic rows
        out.loc[dyn_mask, col] = ( (dyn_end_x - sx) ** 2 + (dyn_end_y - sy) ** 2 ) ** 0.5

    return out


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--input-xlsx", required=True, help="Excel file with dynamic rows and min_dist_line_px/min_dist_tail_px")
    p.add_argument("--csv-dir", required=True, help="Directory containing per-slide CSVs (named <slide_id>.csv)")
    p.add_argument("--output-xlsx", default="Preferences_with_static_distances.xlsx", help="Output Excel file")
    args = p.parse_args()

    input_xlsx = Path(args.input_xlsx)
    csv_dir = Path(args.csv_dir)

    xls = pd.ExcelFile(input_xlsx)

    with pd.ExcelWriter(args.output_xlsx, engine="openpyxl") as w:
        for sheet in xls.sheet_names:
            df = pd.read_excel(input_xlsx, sheet_name=sheet)

            if df.empty:
                df.to_excel(w, sheet_name=sheet[:31], index=False)
                continue

            if "slide_id" not in df.columns:
                raise ValueError(f"Sheet '{sheet}' missing 'slide_id' column")

            slide_id = str(df["slide_id"].iloc[0])
            static_df = load_static_csv(csv_dir, slide_id)

            df2 = add_static_distance_columns(df, static_df)
            df2.to_excel(w, sheet_name=sheet[:31], index=False)

            print(f" {sheet}: added {len(static_df)} static-distance columns from {slide_id}.csv")

    print(f"\nDONE Wrote: {args.output_xlsx}")


if __name__ == "__main__":
    main()

#############
# HOW TO RUN:
#   python3 add_static_distances_to_excel.py \
#     --input-xlsx motion_intent_analysis.xlsx \
#     --csv-dir csv_location_and_distances \
#     --output-xlsx Preferences_with_static_distances.xlsx
#
#############
