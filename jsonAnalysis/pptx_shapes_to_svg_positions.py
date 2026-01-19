#!/usr/bin/env python3

"""
Export shape positions to CSV as SVG-space (px, 96 DPI) coordinates.
Output mode (Option A): one CSV per PPTX file.
Supports both: --pptx single file OR --input-dir folder of many PPTX files.
"""

import argparse
import os
import math
from pathlib import Path
from typing import Tuple, List, Dict

import numpy as np
import pandas as pd
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.exc import PackageNotFoundError

EMU_PER_IN = 914400.0
DPI = 96.0
EMU_PER_PX = EMU_PER_IN / DPI  # 9525.0


def emu_to_px(v_emu: float) -> float:
    return float(v_emu) / EMU_PER_PX


def svg_matrix_from_shape(shape, bbox_emu: Tuple[float, float] = None) -> np.ndarray:
    xfrm = shape.element.xfrm
    if xfrm is None:
        return np.identity(3)

    def val_or_0(v):
        return emu_to_px(v) if v is not None else 0.0

    Bx = val_or_0(xfrm.chOff.x) if xfrm.chOff is not None else 0.0
    By = val_or_0(xfrm.chOff.y) if xfrm.chOff is not None else 0.0
    Dx = val_or_0(xfrm.chExt.cx) if xfrm.chExt is not None else emu_to_px(bbox_emu[0] if bbox_emu else 0.0)
    Dy = val_or_0(xfrm.chExt.cy) if xfrm.chExt is not None else emu_to_px(bbox_emu[1] if bbox_emu else 0.0)

    Bx_ = val_or_0(xfrm.off.x)
    By_ = val_or_0(xfrm.off.y)
    Dx_ = val_or_0(xfrm.ext.cx)
    Dy_ = val_or_0(xfrm.ext.cy)

    theta = (xfrm.rot or 0) * math.pi / 180.0
    Fx = -1.0 if xfrm.flipH else 1.0
    Fy = -1.0 if xfrm.flipV else 1.0

    sx = Dx_ / Dx if Dx != 0 else 1.0
    sy = Dy_ / Dy if Dy != 0 else 1.0
    tx = Bx_ - sx * Bx
    ty = By_ - sy * By
    T_st = np.array([[sx, 0, tx], [0, sy, ty], [0, 0, 1]], dtype=float)

    cx, cy = Bx_ + Dx_ / 2.0, By_ + Dy_ / 2.0
    U = np.array([[1, 0, -cx], [0, 1, -cy], [0, 0, 1]], dtype=float)
    R = np.array([[math.cos(theta), -math.sin(theta), 0],
                  [math.sin(theta), math.cos(theta), 0],
                  [0, 0, 1]], dtype=float)
    Flip = np.array([[Fx, 0, 0], [0, Fy, 0], [0, 0, 1]], dtype=float)
    U_inv = np.array([[1, 0, cx], [0, 1, cy], [0, 0, 1]], dtype=float)
    T_rf = U_inv @ R @ Flip @ U
    return T_rf @ T_st


def apply_matrix(M: np.ndarray, x: float, y: float) -> Tuple[float, float]:
    a, c, e = M[0]
    b, d, f = M[1]
    return a * x + c * y + e, b * x + d * y + f


def shape_type_name(shape) -> str:
    try:
        return MSO_SHAPE_TYPE(shape.shape_type).name
    except Exception:
        return str(shape.shape_type)


def traverse_shapes(shapes, parent_M: np.ndarray, rows: List[Dict], slide_index: int, pptx_file: str):
    for shp in shapes:
        if shp.shape_type == MSO_SHAPE_TYPE.GROUP:
            M_group = svg_matrix_from_shape(shp, bbox_emu=(shp.width, shp.height))
            M_accum = parent_M @ M_group
            traverse_shapes(shp.shapes, M_accum, rows, slide_index, pptx_file)
            continue

        try:
            M_self = svg_matrix_from_shape(shp, bbox_emu=(shp.width, shp.height))
        except Exception:
            M_self = np.identity(3)

        M_final = parent_M @ M_self
        bbox_w_px, bbox_h_px = emu_to_px(shp.width), emu_to_px(shp.height)

        corners = [(0, 0), (bbox_w_px, 0), (0, bbox_h_px), (bbox_w_px, bbox_h_px)]
        pts = [apply_matrix(M_final, x, y) for (x, y) in corners]
        xs, ys = [p[0] for p in pts], [p[1] for p in pts]
        minx, maxx, miny, maxy = min(xs), max(xs), min(ys), max(ys)
        a, c, e = M_final[0]
        b, d, f = M_final[1]
        rotation_deg = math.degrees(math.atan2(b, a))

        rows.append({
            "pptx_file": pptx_file,
            "slide_index": slide_index,
            "shape_name": getattr(shp, "name", ""),
            "shape_type": shape_type_name(shp),
            "left_px": round(minx, 4),
            "top_px": round(miny, 4),
            "width_px": round(maxx - minx, 4),
            "height_px": round(maxy - miny, 4),
            "center_x_px": round((minx + maxx) / 2.0, 4),
            "center_y_px": round((miny + maxy) / 2.0, 4),
            "matrix_a": a, "matrix_b": b, "matrix_c": c,
            "matrix_d": d, "matrix_e": e, "matrix_f": f,
            "rotation_deg_est": round(rotation_deg, 4),
        })


def process_pptx(pptx_path: str) -> pd.DataFrame:
    prs = Presentation(pptx_path)
    rows_all = []
    for si, slide in enumerate(prs.slides, start=1):
        rows: List[Dict] = []
        traverse_shapes(slide.shapes, np.identity(3), rows, si, os.path.basename(pptx_path))
        rows_all.extend(rows)
    return pd.DataFrame(rows_all)


def collect_pptx_files(input_dir: str) -> List[str]:
    pptx_files = []
    for f in os.listdir(input_dir):
        full = os.path.join(input_dir, f)
        if os.path.isfile(full) and f.lower().endswith(".pptx") and not f.startswith("~$"):
            pptx_files.append(full)
    return pptx_files


def main():
    ap = argparse.ArgumentParser(description="Export PPTX shape positions to CSV (one CSV per PPTX)")
    ap.add_argument("--pptx", help="Single PPTX file")
    ap.add_argument("--input-dir", help="Folder containing PPTX files")
    ap.add_argument("--output-dir", default="pptx_shapes_positions_csv", help="Output directory for CSV files")
    args = ap.parse_args()

    if not args.pptx and not args.input_dir:
        ap.error("Please provide either --pptx or --input-dir")

    pptx_files = []
    if args.pptx:
        if args.pptx.startswith("~$"):
            raise SystemExit("That looks like a temporary lock file (~$...). Use the real PPTX instead.")
        pptx_files = [args.pptx]
    else:
        pptx_files = collect_pptx_files(args.input_dir)

    if not pptx_files:
        raise SystemExit("No .pptx files found")

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    for pptx_path in sorted(pptx_files):
        print(f"Processing: {os.path.basename(pptx_path)}")
        try:
            df = process_pptx(pptx_path)
        except PackageNotFoundError as e:
            print(f"Skipping invalid PPTX ({pptx_path}): {e}")
            continue

        out_csv = output_dir / (Path(pptx_path).stem + ".csv")
        df.to_csv(out_csv, index=False, encoding="utf-8")
        print(f"  -> {out_csv}")

    print(f"\nSaved CSV files to: {output_dir}")


if __name__ == "__main__":
    main()


#############
# HOW TO RUN:
#
#   1) Single PPTX file:
#       python3 pptx_shapes_to_svg_positions.py \
#           --pptx /path/to/presentation.pptx \
#           --output-dir csv_shapes
#
#   2) A folder containing many PPTX files:
#       python3 pptx_shapes_to_svg_positions.py \
#           --input-dir /path/to/folder_with_pptx \
#           --output-dir csv_shapes
#
# Explanation of Arguments:
#   --pptx /path/to/presentation.pptx
#       Path to a single PPTX file to process.
#       Use this OR --input-dir (exactly one must be provided).
#
#   --input-dir /path/to/folder
#       Path to a folder containing multiple .pptx files. The script will scan
#       the folder and process every valid PPTX (skipping "~$..." temp files).
#       Use this OR --pptx (not both).
#
#   --output-dir
#       Output directory where CSV files will be written (default: pptx_shapes_positions_csv).
#       Each processed PPTX gets its own CSV file (file name = PPTX name + .csv).
#
# Notes:
#   - Coordinates are converted to SVG-space pixels at 96 DPI
#     (1 px = 9525 EMU in PowerPoint units).
#   - Grouped shapes are handled recursively: all parent transforms
#     (scale, rotation, flip, translate) are applied to child shapes.
#
# OUTPUT:
#   - One CSV per PPTX file.
#   - Each row is a shape on a slide, with coordinates in SVG pixel units (96 DPI).
#   - Columns include:
#       pptx_file         — the source PPTX filename
#       slide_index       — slide number
#       shape_name        — shape name (from PowerPoint)
#       shape_type        — PowerPoint shape type (e.g., AUTO_SHAPE, GROUP, etc.)
#       left_px, top_px  – top-left of the shape’s axis-aligned bounding box (AABB)
#       width_px, height_px – AABB dimensions (AABB - Axis-Aligned Bounding Box)
#       center_x_px, center_y_px – center of the AABB (AABB - Axis-Aligned Bounding Box)
#       matrix_a..matrix_f – SVG transform matrix components (a,b,c,d,e,f)
#       rotation_deg_est – estimated rotation (degrees) derived from the matrix
#############
