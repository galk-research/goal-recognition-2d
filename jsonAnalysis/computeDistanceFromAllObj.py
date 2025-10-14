#!/usr/bin/env python3

"""
Compute distances between the dynamic object and all other objects in each slide.

For every slide in a folder of JSON files:
- Identifies the dynamic objects motion path (from the JSON data).
- Calculates two types of distances between the dynamic object and
  every other object (both static and dynamic) on the same slide:
    * Optimal distance is the straight-line distance from the dynamic
      object's start point to the center of the target object.
    * Real distance is the actual path length traveled by the dynamic
      object along its motion trajectory.

Outputs:
- One Excel sheet per slide.
- Each row represents the distance between the dynamic object and
  one object on that slide (static or dynamic).
- Includes coordinates, path details, and source information
  (whether taken from Excel or JSON).

Requires:
    pip install pandas openpyxl
"""

import argparse
import json
import math
import os
import re
from glob import glob
from typing import Dict, Any, List, Tuple, Optional

import pandas as pd


# ==================== Helpers ====================

def _euclid(p1, p2):
    return math.hypot(p2[0] - p1[0], p2[1] - p1[1])


def _nearly_equal(a, b, tol=1e-6):
    return abs(a - b) <= tol


def _point_equal(p, q, tol=1e-6):
    return _nearly_equal(p[0], q[0], tol) and _nearly_equal(p[1], q[1], tol)


# ==================== JSON parsing ====================

def _get_dynamic_point(locs):
    if not isinstance(locs, list):
        return None
    for d in locs:
        try:
            if str(d.get("src")) == "dynamic":
                return (float(d["x"]), float(d["y"]))
        except Exception:
            continue
    return None


def _infer_start_from_moves(moves, rel_moves, init_locations):
    if isinstance(moves, list) and len(moves) > 0:
        try:
            mx0, my0 = float(moves[0]["x"]), float(moves[0]["y"])
            if isinstance(rel_moves, list) and len(rel_moves) > 0:
                rx0, ry0 = float(rel_moves[0]["x"]), float(rel_moves[0]["y"])
                return (mx0 - rx0, my0 - ry0)
            return (mx0, my0)
        except Exception:
            pass
    return _get_dynamic_point(init_locations)


def _collect_path_points(start, moves, end):
    pts = []
    start_included = False
    end_included = False

    if start is not None:
        pts.append(start)
        start_included = True

    if isinstance(moves, list):
        for m in moves:
            try:
                x, y = float(m["x"]), float(m["y"])
                if not pts or not _point_equal((x, y), pts[-1]):
                    pts.append((x, y))
            except Exception:
                continue

    if end is not None:
        if not pts or not _point_equal(end, pts[-1]):
            pts.append(end)
        end_included = True

    return pts, start_included, end_included


def _trim_trailing_zero_endpoint(pts):
    while len(pts) >= 2 and (pts[-1][0] == 0 or pts[-1][1] == 0):
        pts.pop()
    return pts


def _axis_path_length(pts, axis):
    if len(pts) < 2:
        return 0.0
    return sum(abs(pts[i][axis] - pts[i - 1][axis]) for i in range(1, len(pts)))


# ==================== Excel Statics Loader ====================

def _load_statics_from_excel(excel_path, pptx_name, exclude_names, debug=False):
   
    xls = pd.ExcelFile(excel_path)
    dfs = [pd.read_excel(excel_path, sheet_name=s) for s in xls.sheet_names]
    df_all = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

    required = {"pptx_file", "slide_index", "shape_name", "center_x_px", "center_y_px"}
    missing = required - set(map(str, df_all.columns))
    if missing:
        raise ValueError(f"Excel missing columns: {missing}. Expected {required}")

    def _norm(s):
        return re.sub(r"\s+", "", str(s).strip().lower())

    if pptx_name:
        q = _norm(pptx_name)
        q_noext = _norm(os.path.splitext(pptx_name)[0])
        def _match(cell):
            c = _norm(cell)
            return c == q or c == q_noext or c.endswith(q) or c.endswith(q_noext)
        before = len(df_all)
        df_all = df_all[df_all["pptx_file"].apply(_match)]
        if debug:
            uniq = sorted(df_all["pptx_file"].astype(str).unique().tolist())
            print(f"[DEBUG] Filter by pptx_name='{pptx_name}': {before}->{len(df_all)} rows. Matches: {uniq}")
    else:
        if debug:
            print("[DEBUG] Using ALL rows (no pptx-name filter).")
            print("[DEBUG] pptx_file values:", sorted(df_all["pptx_file"].astype(str).unique().tolist()))

    exclude_norm = {_norm(s) for s in exclude_names if str(s).strip()}
    out = {}
    for _, r in df_all.iterrows():
        try:
            slide_idx = int(r["slide_index"])
            name = str(r["shape_name"]).strip()
            cx, cy = float(r["center_x_px"]), float(r["center_y_px"])
        except Exception:
            continue
        if _norm(name) in exclude_norm:
            continue
        out.setdefault(slide_idx, []).append((name, (cx, cy)))

    if debug:
        total = sum(len(v) for v in out.values())
        print(f"[DEBUG] Loaded statics for {len(out)} slides (total shapes: {total})")
    return out


def _slide_index_from_id(slide_id):
    if not slide_id:
        return None
    s = str(slide_id)
    m = re.search(r"slide[\s_\-]*([0-9]+)", s, flags=re.IGNORECASE)
    if m:
        return int(m.group(1))
    m2 = re.findall(r"([0-9]+)", s)
    if m2:
        return int(m2[-1])
    return None


def _list_static_points_from_json(final_locs, init_locs):
    out = []
    if not isinstance(final_locs, list):
        return out
    for d in final_locs:
        try:
            if str(d.get("src")) != "dynamic":
                out.append(("unknown", (float(d["x"]), float(d["y"]))))
        except Exception:
            continue
    return out


# ==================== Main Processor ====================

def process_folder(folder, excel_path=None, pptx_name=None, exclude_shapes=None, debug=False):
    exclude_names = []
    if exclude_shapes:
        exclude_names = [s for s in exclude_shapes.split(";") if s.strip()]
    if not exclude_names:
        exclude_names = ["Oval 8"]

    statics_by_slideindex = {}
    if excel_path:
        statics_by_slideindex = _load_statics_from_excel(excel_path, pptx_name, exclude_names, debug)

    per_slide_rows = {}
    json_files = sorted(glob(os.path.join(folder, "*.json")))

    if debug:
        print(f"[DEBUG] Found {len(json_files)} JSON files in folder '{folder}'")

    for jf in json_files:
        try:
            with open(jf, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            print(f"[WARN] Skipping {jf}: {e}")
            continue

        for slide_id, slide_blob in data.items():
            slide_index = _slide_index_from_id(slide_id)
            if debug and slide_index is None:
                print(f"[DEBUG] Could not extract slide_index from slide_id='{slide_id}'")

            for worker_id, worker_blob in slide_blob.items():
                for assignment_id, assignment in worker_blob.items():
                    init_locs = assignment.get("init_locations", [])
                    final_locs = assignment.get("final_locations", [])
                    moves = assignment.get("moves", [])
                    rel_moves = assignment.get("relative moves", assignment.get("relative_moves", []))

                    start = _infer_start_from_moves(moves, rel_moves, init_locs)
                    dyn_end = _get_dynamic_point(final_locs)

                    pts, start_included, end_included = _collect_path_points(start, moves, dyn_end)
                    pts = _trim_trailing_zero_endpoint(pts)
                    if pts:
                        start = pts[0]

                    if len(pts) >= 2:
                        real_2d_full = sum(_euclid(pts[i - 1], pts[i]) for i in range(1, len(pts)))
                        real_x_full = _axis_path_length(pts, 0)
                        real_y_full = _axis_path_length(pts, 1)
                    else:
                        real_2d_full = real_x_full = real_y_full = 0.0

                    endpoints = []
                    if dyn_end is not None:
                        endpoints.append(("dynamic", dyn_end))

                    statics = []
                    if slide_index and slide_index in statics_by_slideindex:
                        statics = statics_by_slideindex[slide_index]
                        statics = [(n, p) for n, p in statics if p is not None]
                    else:
                        statics = _list_static_points_from_json(final_locs, init_locs)

                    endpoints.extend(statics)

                    for obj_name, (ex, ey) in endpoints:
                        if start is not None:
                            optimal_2d = _euclid(start, (ex, ey))
                            optimal_x = abs(ex - start[0])
                            optimal_y = abs(ey - start[1])
                        else:
                            optimal_2d = optimal_x = optimal_y = float("nan")

                        row = {
                            "slide_id": slide_id,
                            "slide_index": slide_index,
                            "worker_id": worker_id,
                            "assignment_id": assignment_id,
                            "object_name": obj_name,
                            "start_x_dinamic_obj": start[0] if start else float("nan"),
                            "start_y_dinamic_obj": start[1] if start else float("nan"),
                            "end_x": ex,
                            "end_y": ey,
                            "moves_count": len(moves) if isinstance(moves, list) else 0,
                            "optimal_2d": optimal_2d,
                            "real_2d": real_2d_full,
                            "optimal_x": optimal_x,
                            "real_x": real_x_full,
                            "optimal_y": optimal_y,
                            "real_y": real_y_full,
                        }
                        per_slide_rows.setdefault(slide_id, []).append(row)

    per_slide = {k: pd.DataFrame(v) for k, v in per_slide_rows.items()}
    all_rows = pd.concat(per_slide.values(), ignore_index=True) if per_slide else pd.DataFrame()
    return all_rows, per_slide


# ==================== Excel Writer ====================

def write_excel(out_path, all_rows_df, per_slide):
    if all_rows_df.empty:
        print("[WARN] No data to write.")
        return
    with pd.ExcelWriter(out_path, engine="openpyxl") as xl:
        for slide_id, df in per_slide.items():
            df_out = df.copy()
            desired = [
                "slide_id", "slide_index", "worker_id", "assignment_id",
                "object_name",
                "moves_count",
                "start_x_dinamic_obj", "start_y_dinamic_obj", "end_x", "end_y",
                "optimal_2d", "real_2d",
                "optimal_x", "real_x",
                "optimal_y", "real_y",
            ]
            cols = [c for c in desired if c in df_out.columns]
            df_out.sort_values(["worker_id", "assignment_id", "object_name"], inplace=True)
            sheet_name = (str(slide_id)[:28] or "slide")
            df_out.to_excel(xl, sheet_name=sheet_name, index=False, columns=cols)


# ==================== main ====================

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("folder", help="Folder containing slide JSON files")
    parser.add_argument("-o", "--output", default="distancesFromAllObj.xlsx")
    parser.add_argument("--excel", help="Excel file with shape centers")
    parser.add_argument("--pptx-name", help="Filter pptx_file in Excel")
    parser.add_argument("--exclude-shapes", default="Oval 8")
    parser.add_argument("--debug", action="store_true")
    args = parser.parse_args()

    all_df, per_slide = process_folder(
        args.folder,
        excel_path=args.excel,
        pptx_name=args.pptx_name,
        exclude_shapes=args.exclude_shapes,
        debug=args.debug,
    )
    write_excel(args.output, all_df, per_slide)
    print(f"Done. Wrote: {args.output}")


if __name__ == "__main__":
    main()


#############
# HOW TO RUN:
#
#   run - python3 name_of_file.py /path/to/json_folder \
#      -o name_of_file.xlsx \
#      --excel obj_positions.xlsx \
#      --pptx-name "group 8 - Similarity + Continuation.pptx" \
#      --exclude-shapes "Oval 8" 
#
# Explanation of Arguments:
#   /path/to/json_folder
#       Path to the folder containing all slide-JSON files.
#       The script will process every *.json file inside this folder.
#
#   -o name_of_file.xlsx   (or --output name_of_file.xlsx)
#       Name (or full path) of the Excel file to create.
#
#   --excel obj_positions.xlsx
#       Path to the Excel file generated from PowerPoint
#       (via pptx_shapes_to_svg_positions.py).
#       This file provides the center positions (in SVG pixel units)
#       of all static shapes on each slide.
#
#   --pptx-name "group 8 - Similarity + Continuation.pptx"
#       Filters rows from the Excel file so that only shapes belonging
#       to this specific PowerPoint presentation are used.
#       The filter is lenient: it works with or without the “.pptx”
#       extension and ignores spaces/case differences.
#       If you only have one presentation in the Excel file,
#       you can safely omit this argument.
#
#   --exclude-shapes "Oval 8"
#       Names of shapes to exclude from the static set.
#       Multiple names can be separated by semicolons, e.g.
#       The match is case-insensitive and ignores extra spaces.
#       Oval 8 is the dynamic location before the animation starts.
#
#   --debug
#       Prints detailed log information during execution:
#       number of JSON files found, slides matched, shapes loaded
#       from Excel, and other helpful diagnostics.
#       Recommended for the first run.
#
# OUTPUT:
#   - One Excel sheet per slide.
#   - Each row represents the distance between the dynamic object
#     and one shape on that slide (static or dynamic).
#   - Columns include:
#       object_name — "dynamic" or the static shape name
#       start_x, start_y — dynamic object’s start point (from JSON)
#       end_x, end_y — target shape center (from Excel or JSON)
#       optimal_* — straight-line distance from start→end
#       real_* — total path length actually traveled by the dynamic object

#############
