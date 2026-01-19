# JSON Analysis
This repository contains a full end-to-end pipeline for analyzing motion intentionality and preferences in PPTX-based visual stimuli, combining JSON logs, PowerPoint geometry, and motion-line analysis.

The pipeline extracts per-group and per-slide data, computes spatial and motion-based distances, and produces final Excel files with Intentionality and Preferences metrics.

---

## Recommended Folder Structure

```text
project_root/
│
├── raw/
│   ├── all_groups.json
│   └── pptx/
│       └── group NN - Similarity + Continuation.pptx
│
├── groups/
│   └── groupNN.json
│
├── slides/
│   └── group0x/
│       ├── svg-groupNN_slide01.json
│       ├── svg-groupNN_slide02.json
│       └── ...
│
├── shapes/
│   └── pptx_shapes_positions.xlsx
│
├── lines/
│   └── motion_lines_metrics_csv/
│       ├── svg-groupNN_slide01.csv
│       └── ...
│
├── distances/
│   └── csv_location_and_distances/
│       ├── svg-groupNN_slide01.csv
│       └── ...
│
├── results/
│   ├── motion_intent_analysis.xlsx
│   ├── Preferences_with_static_distances.xlsx
│   └── Intentionality.xlsx
│
└── scripts/
    └── *.py
```

---

## Pipeline Overview

| Step | Script                                 | Purpose                                                     |
| ---- | -------------------------------------- | ----------------------------------------------------------- |
| 1    | extract_json_by_group.py               | Extract one group from the large JSON                       |
| 2    | split_json_by_slide.py                 | Split group JSON into per-slide JSONs                       |
| 3    | pptx_shapes_to_svg_positions.py        | Export PPTX shape geometry                                  |
| 4    | extract_motion_lines.py                | Extract motion line metrics                                 |
| 5    | computeDistanceFromAllObj.py           | Compute distances from dynamic object to all static objects |
| 6    | filter_dynamic_rows.py                 | Aggregate dynamic rows                                      |
| 7    | add_line_distances_to_intent.py        | Add distance-to-line metrics                                |
| 8    | add_static_distances_to_preferences.py | Add static object distances                                 |
| 9    | build_intentionality_from_scratch.py   | Compute Intentionality metrics                              |
| 10   | run_full_pipeline.py                   | Run everything end-to-end                                   |

---

## Scripts

### extract_json_by_group.py

This script extracts all slides of a specific group from a large JSON file. The result is a smaller JSON that contains only the slides of that group.

**How to run**

```bash
python3 extract_json_by_group.py <group_number> \
    --input_file <input_file_path> \
    --output_dir <output_directory>
```

---

### split_json_by_slide.py

This script takes a group JSON (for one group only) and splits it into one JSON file per slide.

For each key starting with `svg-groupNN_slide`, the script:

* Creates a separate JSON containing only that key.
* Saves it under a group-specific folder, e.g. `group03/`.

**How to run**

```bash
python3 split_json_by_slide.py <group_number> \
    --input_file <input_file_path> \
    --output_dir <output_directory>
```

---

### pptx_shapes_to_svg_positions.py

This script reads PowerPoint (.pptx) files and exports shape positions into an CSV files.

For each shape on each slide, it computes:

* Position and size in pixels at 96 DPI (converted from EMU).
* Center coordinates (center_x_px, center_y_px).
* A 2D transform matrix (matrix_a … matrix_f).
* An estimated rotation angle (rotation_deg_est).

It supports:

* One PPTX file (with `--pptx`)
* Or a folder of PPTX files (with `--input-dir`)

Each PPTX becomes a separate sheet in the output Excel.

**How to run**

Option 1 – Single PPTX

```bash
python3 pptx_shapes_to_svg_positions.py \
    --pptx /path/to/presentation.pptx \
    --output-dir csv_shapes
```

Option 2 – Folder of PPTX files

```bash
python3 pptx_shapes_to_svg_positions.py \
    --input-dir /path/to/pptx_folder \
    --output-dir csv_shapes
```

---

### extract_motion_line.py

Extract motion-line information from PPTX files and compute two line metrics:

* **main line** – straight line from shape's connector endpoints (if available), otherwise fitted (via PCA) to ALL path points, otherwise bbox-based.
* **tail line** – straight line defined by the LAST TWO path points.

**How to run**

Option 1 – Single PPTX

```bash
python3 extract_motion_lines.py \
    /full/path/to/presentation.pptx \
    --output-dir motion_lines_metrics_csv
```

Option 2 – Folder of PPTX files

```bash
python3 extract_motion_lines.py \
    /full/path/to/pptx_folder \
    --output-dir motion_lines_metrics_csv
```

---

### computeDistanceFromAllObj.py

Compute distances between the dynamic object and all other objects in each slide.

For every slide in a folder of JSON files:

* Identifies the dynamic object's motion path (from the JSON data).
* Calculates two types of distances between the dynamic object and every other object (static + dynamic) on the same slide:

  * **Optimal distance**: straight-line distance from the dynamic start point to the target object's center.
  * **Real distance**: the actual path length traveled by the dynamic object along its motion trajectory (sum of segment distances).

**How to run**

```bash
python3 computeDistanceFromAllObj.py /path/to/json_folder \
    -o csv_location_and_distances \
    --centers-csv /path/to/pptx_shapes_positions_csv \
    --pptx-name "group NN - Similarity + Continuation.pptx" \
    --exclude-shapes "Oval 8"
```

---

### filter_dynamic_rows.py

Filter rows related to the dynamic object from distance CSV files and aggregate them into a single Excel file.

**How to run**

```bash
python3 filter_dynamic_rows.py \
    /full/path/to/distances_csv \
    --output-file motion_intent_analysis.xlsx
```

---

### add_line_distances_to_intent.py

Add two distance-to-line columns to `motion_intent_analysis.xlsx`:

* `min_dist_line_px` : minimal distance from the dynamic final point (end_x, end_y) to any MAIN line in the same slide.
* `min_dist_tail_px` : minimal distance from the dynamic final point (end_x, end_y) to any TAIL line in the same slide (only if tail_points_used > 0).

**How to run**

```bash
python3 add_line_distances_to_intent.py \
    --intent-file motion_intent_analysis.xlsx \
    --lines-file motion_lines_metrics_csv/
```

---

### add_static_distances_to_preferences.py

Add per-static distance columns to an existing Excel file.

**How to run**

```bash
python3 add_static_distances_to_excel.py \
    --input-xlsx motion_intent_analysis.xlsx \
    --csv-dir csv_location_and_distances \
    --output-xlsx Preferences_with_static_distances.xlsx
```

---

### build_intentionality_from_scratch.py

Compute Intentionality metrics toward static objects and write them to Excel.

**How to run**

```bash
python3 build_intentionality_from_csv.py \
    --input-xlsx motion_intent_analysis.xlsx \
    --csv-dir csv_location_and_distances \
    --output-xlsx Intentionality.xlsx
```

---

### run_full_pipeline.py

Run the full pipeline end-to-end.

**How to run**

```bash
python3 run_full_pipeline.py
```

---

## Output

The output is a merge of the two Excel files: **Preferences** and **Intentionality**.
