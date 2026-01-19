#!/usr/bin/env python3
"""
Add two columns to motion_intent_analysis.xlsx:

- min_dist_line_px : minimal distance from dynamic final point to any MAIN line
- min_dist_tail_px : minimal distance to any TAIL line

INPUTS:
1) motion_intent_analysis.xlsx
2) motion_lines_metrics CSV:
   - a single CSV file OR
   - a directory containing multiple CSV files (one per PPTX)

OUTPUT (Excel):
- Excel file containing all original sheets + 2 new columns:
    min_dist_line_px, min_dist_tail_px

Requires:
    pip install pandas openpyxl numpy
"""



import argparse
from pathlib import Path
import re
from typing import Dict

import numpy as np
import pandas as pd


# ---------- Geometry ----------

def point_to_line_distance(px, py, x1, y1, x2, y2):
    if any(np.isnan([px, py, x1, y1, x2, y2])):
        return np.nan
    if x1 == x2 and y1 == y2:
        return float(np.hypot(px - x1, py - y1))
    vx, vy = x2 - x1, y2 - y1
    num = abs(vy * px - vx * py + x2 * y1 - y2 * x1)
    den = (vx * vx + vy * vy) ** 0.5
    return float(num / den)


def slide_index_from_slide_id(slide_id: str):
    if not isinstance(slide_id, str):
        return np.nan
    m = re.search(r"slide(\d+)", slide_id)
    return float(m.group(1)) if m else np.nan


# ---------- CSV loader (replaces Excel) ----------

def load_lines_csv(path: Path) -> pd.DataFrame:
    if path.is_file():
        if path.suffix.lower() != ".csv":
            raise SystemExit("lines-file must be a CSV or a directory of CSVs")
        dfs = [pd.read_csv(path)]
    else:
        dfs = [pd.read_csv(p) for p in sorted(path.glob("*.csv"))]

    if not dfs:
        raise SystemExit("No CSV files found for motion lines")

    df = pd.concat(dfs, ignore_index=True)

    required = [
        "slide_index",
        "line_x1_px", "line_y1_px",
        "line_x2_px", "line_y2_px",
        "tail_x1_px", "tail_y1_px",
        "tail_x2_px", "tail_y2_px",
        "tail_points_used",
    ]
    for col in required:
        if col not in df.columns:
            raise SystemExit(f"Column '{col}' missing from motion_lines CSV")

    return df


# ---------- main ----------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--intent-file", required=True, help="motion_intent_analysis.xlsx")
    ap.add_argument(
        "--lines-file",
        required=True,
        help="CSV file OR directory with motion_lines_metrics CSVs",
    )
    ap.add_argument(
        "--output-file",
        default=None,
        help="Output Excel file (default: intent-file with _with_distances suffix)",
    )
    ap.add_argument("--intent-x-col", default="end_x")
    ap.add_argument("--intent-y-col", default="end_y")
    args = ap.parse_args()

    intent_path = Path(args.intent_file).resolve()


    if args.output_file:
        output_path = Path(args.output_file).resolve()
    else:
        output_path = intent_path
    lines_path = Path(args.lines_file).resolve()

    # ----- Load motion lines (CSV) -----
    lines_df = load_lines_csv(lines_path)

    lines_by_slide: Dict[float, pd.DataFrame] = {
        idx: g.reset_index(drop=True)
        for idx, g in lines_df.groupby("slide_index")
    }

    # ----- Load intent Excel (all sheets) -----
    sheets = pd.read_excel(intent_path, sheet_name=None)
    updated = {}

    for sheet_name, df in sheets.items():
        df = df.copy()

        if "slide_id" not in df.columns:
            updated[sheet_name] = df
            continue

        if args.intent_x_col not in df.columns or args.intent_y_col not in df.columns:
            updated[sheet_name] = df
            continue

        if "slide_index" not in df.columns:
            df["slide_index"] = df["slide_id"].astype(str).apply(slide_index_from_slide_id)

        min_line, min_tail = [], []

        for _, r in df.iterrows():
            px, py = float(r[args.intent_x_col]), float(r[args.intent_y_col])
            slide_idx = r["slide_index"]

            best_line = np.nan
            best_tail = np.nan

            lines = lines_by_slide.get(slide_idx)
            if lines is not None:
                line_d, tail_d = [], []
                for _, l in lines.iterrows():
                    d_line = point_to_line_distance(
                        px, py,
                        l["line_x1_px"], l["line_y1_px"],
                        l["line_x2_px"], l["line_y2_px"],
                    )
                    if not np.isnan(d_line):
                        line_d.append(d_line)

                    if l["tail_points_used"] > 0:
                        d_tail = point_to_line_distance(
                            px, py,
                            l["tail_x1_px"], l["tail_y1_px"],
                            l["tail_x2_px"], l["tail_y2_px"],
                        )
                        if not np.isnan(d_tail):
                            tail_d.append(d_tail)

                if line_d:
                    best_line = min(line_d)
                if tail_d:
                    best_tail = min(tail_d)

            min_line.append(best_line)
            min_tail.append(best_tail)

        df["min_dist_line_px"] = min_line
        df["min_dist_tail_px"] = min_tail
        updated[sheet_name] = df


    with pd.ExcelWriter(output_path, engine="openpyxl") as w:
        for sheet, df in updated.items():
            df.to_excel(w, sheet_name=sheet, index=False)

    print(f"Updated Excel file:\n  {output_path}")
    print("Added/updated columns: min_dist_line_px, min_dist_tail_px")



if __name__ == "__main__":
    main()

#############
# HOW TO RUN:
#
# python3 add_line_distances_to_intent.py \
# --intent-file motion_intent_analysis.xlsx \
# --lines-file motion_lines_metrics_csv/
#
#############
