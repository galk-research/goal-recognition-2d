#!/usr/bin/env python3
"""
Compute Intentionality metrics toward static objects and write them to Excel.

This script takes:
1) An Excel file containing ONLY dynamic-object rows
2) A directory containing per-slide CSV files
   (each CSV named <slide_id>.csv), which include:
   - dynamic START coordinates
   - static object END coordinates

For each Excel sheet (one slide):
- Keeps only dynamic rows.
- Retrieves the dynamic START point from the corresponding slide CSV.
- Identifies all static objects on that slide.
- Computes:
    * general_intentionality = optimal_2d / real_2d
    * per-static intentionality:
        D(START, static) / (D(actual) + D(USER-END, static))

INPUT:
- --input-xlsx : Excel file with dynamic rows only
- --csv-dir    : Directory containing per-slide CSV files (<slide_id>.csv)

OUTPUT:
- Excel file (default: Intentionality.xlsx)
- One sheet per slide + an ERRORS sheet

Notes:
- START coordinates are taken from CSVs (Excel does not contain them).
- Matching between Excel and CSV is done using:
    worker_id + assignment_id
"""

import re
import math
import argparse
from pathlib import Path
import pandas as pd
from typing import List, Optional


def euclid(x1, y1, x2, y2):
    return math.hypot(x2 - x1, y2 - y1)


def safe_div(n, d):
    if d == 0 or pd.isna(d):
        return float("nan")
    return n / d


def sanitize(name: str) -> str:
    s = str(name).strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^0-9A-Za-z_]", "_", s)
    s = re.sub(r"_+", "_", s)
    return s[:80]


def pick_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    for c in candidates:
        if c in df.columns:
            return c
    return None



def load_slide_csv(csv_dir: Path, slide_id: str) -> pd.DataFrame:
    p = csv_dir / f"{slide_id}.csv"
    if not p.exists():
        raise FileNotFoundError(f"Missing CSV for slide_id={slide_id}: {p}")
    return pd.read_csv(p)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input-xlsx", required=True)
    ap.add_argument("--csv-dir", required=True)
    ap.add_argument("--output-xlsx", default="Intentionality.xlsx")
    args = ap.parse_args()

    input_xlsx = Path(args.input_xlsx)
    csv_dir = Path(args.csv_dir)

    xls = pd.ExcelFile(input_xlsx)

    errors = []
    wrote_any_sheet = False

    with pd.ExcelWriter(args.output_xlsx, engine="openpyxl") as w:
        for sheet in xls.sheet_names:
            try:
                df = pd.read_excel(input_xlsx, sheet_name=sheet)
                if df.empty:
                    continue

                # Required columns in the Excel for dynamic rows
                required_excel = {"slide_id", "object_name", "end_x", "end_y", "optimal_2d", "real_2d"}
                missing = [c for c in required_excel if c not in df.columns]
                if missing:
                    raise ValueError(f"Missing in Excel: {missing}")

                dyn = df[df["object_name"] == "dynamic"].copy()
                if dyn.empty:
                    continue

                slide_id = str(dyn["slide_id"].iloc[0])
                slide_csv = load_slide_csv(csv_dir, slide_id)

                # Find START columns IN CSV (because Excel doesn't have them)
                start_x_col = pick_col(
                    slide_csv,
                    ["start_x_dynamic_obj", "start_x_dynamic", "start_x"]
                )
                start_y_col = pick_col(
                    slide_csv,
                    ["start_y_dynamic_obj", "start_y_dynamic", "start_y"]
                )
                if start_x_col is None or start_y_col is None:
                    raise ValueError(
                        f"CSV for slide_id={slide_id} missing START columns. "
                        f"Looked for start_x_dynamic_obj/start_y_dynamic_obj variants."
                    )

                # Match START from CSV dynamic row(s) to each Excel dynamic row.
                # Best: by worker_id + assignment_id if exists in both.
                join_keys = []
                for k in ["worker_id", "assignment_id"]:
                    if k in dyn.columns and k in slide_csv.columns:
                        join_keys.append(k)

                csv_dyn = slide_csv[slide_csv["object_name"] == "dynamic"].copy()
                if csv_dyn.empty:
                    raise ValueError(f"No dynamic row in CSV for slide_id={slide_id} to take START from.")

                if join_keys:
                    csv_dyn = (
                        csv_dyn[join_keys + [start_x_col, start_y_col]]
                        .drop_duplicates(subset=join_keys, keep="first")
                        .rename(columns={start_x_col: "_start_x", start_y_col: "_start_y"})
                    )
                    dyn = dyn.merge(csv_dyn, on=join_keys, how="left")

                    if dyn["_start_x"].isna().any() or dyn["_start_y"].isna().any():
                        sample = dyn[dyn["_start_x"].isna() | dyn["_start_y"].isna()][join_keys].head(10)
                        raise ValueError(
                            f"Failed to merge START for some rows (slide_id={slide_id}). Examples:\n{sample}"
                        )
                else:
                    # Fallback: take the first dynamic START from CSV for all rows
                    dyn["_start_x"] = float(csv_dyn.iloc[0][start_x_col])
                    dyn["_start_y"] = float(csv_dyn.iloc[0][start_y_col])

                # Statics list from CSV
                needed_csv = {"object_name", "end_x", "end_y"}
                if not needed_csv.issubset(set(slide_csv.columns)):
                    raise ValueError(f"CSV missing columns: {sorted(needed_csv - set(slide_csv.columns))}")

                statics = slide_csv[slide_csv["object_name"] != "dynamic"].copy()
                statics = statics.drop_duplicates(subset=["object_name"], keep="first")

                # Base distances (raw)
                
                dyn["general_intentionality"] = dyn.apply(
                    lambda r: safe_div(float(r["optimal_2d"]), float(r["real_2d"])),
                    axis=1
                )

                sx = dyn["_start_x"].astype(float)
                sy = dyn["_start_y"].astype(float)
                ex = dyn["end_x"].astype(float)
                ey = dyn["end_y"].astype(float)
                d_actual = dyn["real_2d"].astype(float)


                # Per static: raw distances + denom + intent
                for _, s in statics.iterrows():
                    name = s["object_name"]
                    suf = sanitize(name)
                    static_x = float(s["end_x"])
                    static_y = float(s["end_y"])

                    d_start = ((sx - static_x) ** 2 + (sy - static_y) ** 2) ** 0.5  # D(START, static)
                    d_end = ((ex - static_x) ** 2 + (ey - static_y) ** 2) ** 0.5    # D(USER-END, static)
                    denom = d_actual + d_end                                        # D(actual) + D(USER-END, static)

                    dyn[f"d_start_to__{suf}"] = d_start
                    dyn[f"d_user_end_to__{suf}"] = d_end
                    dyn[f"d_actual_plus_end_to__{suf}"] = denom
                    dyn[f"intent_to__{suf}"] = [safe_div(n, d) for n, d in zip(d_start, denom)]

                # ---- remove all line-related / preference columns ----
                cols_to_drop = [
                    "min_dist_line_px",
                    "min_dist_tail_px",
                ]

                dyn = dyn.drop(columns=[c for c in cols_to_drop if c in dyn.columns])

                dyn.to_excel(w, sheet_name=sheet[:31], index=False)
                wrote_any_sheet = True
                print(f"{sheet}: wrote Intentionality for slide_id={slide_id} ({len(statics)} statics)")

            except Exception as e:
                errors.append({"sheet": sheet, "error": str(e)})
                print(f"{sheet}: {e}")

        # Always write an ERRORS sheet (and ensure at least one visible sheet exists)
        err_df = pd.DataFrame(errors) if errors else pd.DataFrame([{"sheet": "", "error": ""}])
        err_df.to_excel(w, sheet_name="ERRORS", index=False)
        wrote_any_sheet = True  # ERRORS guarantees workbook is valid

    print(f"\nDONE Wrote: {args.output_xlsx}")


if __name__ == "__main__":
    main()


    #############
# HOW TO RUN:
#
# INPUT:
#   1) Excel file with dynamic-only rows
#      (e.g. motion_intent_analysis.xlsx)
#   2) Directory with per-slide CSV files
#      (one CSV per slide_id, named <slide_id>.csv)
#
# OUTPUT:
#   - A new Excel file containing Intentionality metrics
#   - One sheet per slide + an ERRORS sheet
#
# ------------------------------------------------------------
#
# BASIC USAGE:
#
#   python3 build_intentionality_from_csv.py \
#     --input-xlsx motion_intent_analysis.xlsx \
#     --csv-dir csv_location_and_distances \
#     --output-xlsx Intentionality.xlsx
#
#
#############

