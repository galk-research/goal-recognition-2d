#!/usr/bin/env python3\
import argparse
import json
from typing import Dict, Any, List
import pandas as pd


"""
Create one Excel sheet per worker from a workers-JSON.
Output:
- Excel file with:
  * One sheet per worker_outer_id
  * Rows: worker_outer_id, worker_inner_id, assignment_id, slide_id, age, gender, hand
"""

def read_workers_json(path: str) -> pd.DataFrame:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    rows: List[Dict[str, Any]] = []
    if not isinstance(data, dict):
        raise ValueError("Workers JSON must be an object at the top level.")

    for worker_outer_id, outer_blob in data.items():
        if not isinstance(outer_blob, dict):
            continue

        for worker_inner_id, inner_blob in outer_blob.items():
            if not isinstance(inner_blob, dict):
                continue

            assignment_from_header = inner_blob.get("AssignmentId")
            age = inner_blob.get("age")
            gender = inner_blob.get("gender")
            hand = inner_blob.get("hand")

            hit = inner_blob.get("hit", {})
            if not isinstance(hit, dict):
                continue

            for slide_id, slide_blob in hit.items():
                if not isinstance(slide_blob, dict):
                    continue
                assignment_id = slide_blob.get("AssignmentId", assignment_from_header)

                rows.append(
                    dict(
                        worker_outer_id=str(worker_outer_id),
                        worker_inner_id=str(worker_inner_id),
                        assignment_id=str(assignment_id) if assignment_id is not None else None,
                        slide_id=str(slide_id),
                        age=age,
                        gender=gender,
                        hand=hand,
                    )
                )

    if not rows:
        return pd.DataFrame(
            columns=[
                "worker_outer_id",
                "worker_inner_id",
                "assignment_id",
                "slide_id",
                "age",
                "gender",
                "hand",
            ]
        )

    return pd.DataFrame(rows)


def write_per_worker_sheets(df: pd.DataFrame, out_path: str) -> None:
    sort_cols = [c for c in ["worker_outer_id", "assignment_id", "slide_id"] if c in df.columns]
    df_sorted = df.sort_values(sort_cols, kind="mergesort", na_position="last")

    def safe_name(name: str, used: set) -> str:
        base = (str(name) or "worker")[:31]
        if base not in used:
            used.add(base)
            return base
        for i in range(2, 1000):
            suffix = f"_{i}"
            cand = (str(name)[: 31 - len(suffix)] + suffix) if len(str(name)) >= len(suffix) else f"{base}{suffix}"
            if cand not in used and len(cand) <= 31:
                used.add(cand)
                return cand
        cand = f"{base[:28]}_zzz"
        used.add(cand)
        return cand

    used_names = set()
    with pd.ExcelWriter(out_path, engine="openpyxl") as xl:
        for worker_id, sub in df_sorted.groupby("worker_outer_id", dropna=False):
            sheet = safe_name(str(worker_id), used_names)
            cols = [
                "worker_outer_id",
                "worker_inner_id",
                "assignment_id",
                "slide_id",
                "age",
                "gender",
                "hand",
            ]
            cols = [c for c in cols if c in sub.columns] + [c for c in sub.columns if c not in cols]
            sub[cols].to_excel(xl, sheet_name=sheet, index=False)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workers_json", help="Path to workers JSON")
    ap.add_argument(
        "-o",
        "--output",
        default="workers_data.xlsx",
        help="Output Excel path (default: workers_data.xlsx)",
    )
    args = ap.parse_args()

    df_workers = read_workers_json(args.workers_json)
    write_per_worker_sheets(df_workers, args.output)
    print(f"Done. Wrote: {args.output}")


if __name__ == "__main__":
    main()

################
# HOW TO RUN:
# run - python3 name_of_file.py </path/to/workers.json> -o output.xlsx

# Explanation of Arguments:

# --input_file <input_file_path> (Required):
# Example: --input_file /path/to/your/input_file.json

# --output_dir <output_directory> (Optional):
# This is where the output file will be saved.
# Example: --output_dir /path/to/save/directory/
# If you don't specify this, the output will be saved to the default directory that you provide above.