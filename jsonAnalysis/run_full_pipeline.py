#!/usr/bin/env python3
"""
Run the full pipeline end-to-end.

Asks for:
- Path to BIG JSON
- Path to PPTX file OR folder

Runs in this order:
1) extract_json_by_group.py
2) extract_json_by_slide.py
3) pptx_shapes_to_svg_positions.py
4) extract_motion_lines.py
5) computeDistanceFromAllObj.py
6) filter_dynamic_rows.py
7) add_line_distances_to_intent.py
8) add_static_distances_to_preferences.py
9) build_intentionality_from_scratch.py

Finally:
- MERGES the last two Excel files (Preferences + Intentionality) into ONE,
"""

import argparse
import re
import subprocess
import sys
from pathlib import Path
from typing import List, Optional, Dict, Tuple

import pandas as pd

# ------------------ merge functionality ------------------

def _ensure_str(df: pd.DataFrame, cols):
    df = df.copy()
    for c in cols:
        if c in df.columns:
            df[c] = df[c].astype(str)
    return df


def _best_join_keys(df_left: pd.DataFrame, df_right: pd.DataFrame):
    priority = [
        ["slide_id", "worker_id", "assignment_id"],
        ["slide_id", "assignment_id"],
        ["slide_id", "worker_id"],
        ["slide_id"],
    ]
    for keys in priority:
        if all(k in df_left.columns for k in keys) and all(k in df_right.columns for k in keys):
            return keys
    return []


def merge_intentionality_into_preferences(preferences_xlsx: Path, intentionality_xlsx: Path, out_xlsx: Path):
    """
    Creates ONE Excel where each sheet is the Preferences sheet +
    adds from Intentionality only:
      - general_intentionality
      - columns starting with 'intent_to__'

    Matching strategy (more robust):
    1) Try match by SAME SHEET NAME
    2) Fallback: match by slide_id (normalized)
    """

    pref_sheets = pd.read_excel(preferences_xlsx, sheet_name=None)
    intent_sheets = pd.read_excel(intentionality_xlsx, sheet_name=None)

    def norm_slide_id(x):
        return str(x).strip().lower()

    # fallback map: slide_id -> intent df
    intent_by_slide = {}
    for sh, df in intent_sheets.items():
        if df is None or df.empty or "slide_id" not in df.columns:
            continue
        sid = norm_slide_id(df["slide_id"].iloc[0])
        intent_by_slide[sid] = df

    report = []

    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as w:
        for sh, df_pref in pref_sheets.items():
            if df_pref is None or df_pref.empty:
                df_pref.to_excel(w, sheet_name=sh[:31], index=False)
                report.append({"sheet": sh, "status": "EMPTY_PREF"})
                continue

            # ---- 1) match by sheet name ----
            df_int = intent_sheets.get(sh)

            match_mode = "SHEET_NAME"

            # ---- 2) fallback: match by slide_id ----
            if df_int is None or df_int.empty:
                if "slide_id" in df_pref.columns:
                    sid = norm_slide_id(df_pref["slide_id"].iloc[0])
                    df_int = intent_by_slide.get(sid)
                    match_mode = "SLIDE_ID"
                else:
                    df_int = None

            if df_int is None or df_int.empty:
                df_pref.to_excel(w, sheet_name=sh[:31], index=False)
                report.append({"sheet": sh, "status": "NO_MATCH_IN_INTENT"})
                continue

            # choose keys
            keys = _best_join_keys(df_pref, df_int)
            if not keys:
                df_pref.to_excel(w, sheet_name=sh[:31], index=False)
                report.append({"sheet": sh, "status": "NO_COMMON_KEYS", "match_mode": match_mode})
                continue

            # keep only desired columns from Intentionality
            intent_cols = [
                c for c in df_int.columns
                if c not in keys and (c == "general_intentionality" or str(c).startswith("intent_to__"))
            ]
            if not intent_cols:
                df_pref.to_excel(w, sheet_name=sh[:31], index=False)
                report.append({"sheet": sh, "status": "NO_INTENT_COLS_FOUND", "match_mode": match_mode})
                continue

            # normalize key types to string to avoid int/str merge mismatch
            left = _ensure_str(df_pref, keys)
            right = _ensure_str(df_int[keys + intent_cols], keys)

            # avoid overwriting existing columns
            rename = {c: f"{c}__intent" for c in intent_cols if c in left.columns}
            if rename:
                right = right.rename(columns=rename)

            merged = left.merge(right, on=keys, how="left")
            merged.to_excel(w, sheet_name=sh[:31], index=False)

            report.append({
                "sheet": sh,
                "status": "MERGED",
                "match_mode": match_mode,
                "keys": ",".join(keys),
                "added_cols": len(intent_cols),
            })

        pd.DataFrame(report).to_excel(w, sheet_name="MERGE_REPORT", index=False)




# ------------------ helpers ------------------

def run_step(cmd: List[str], name: str):
    print(f"\n▶ {name}")
    print("  " + " ".join(cmd))
    res = subprocess.run(cmd)
    if res.returncode != 0:
        print(f"\n FAILED: {name}")
        sys.exit(res.returncode)
    print(f"DONE: {name}")


def ask_path(prompt: str) -> Path:
    while True:
        s = input(prompt).strip().strip('"').strip("'")
        p = Path(s).expanduser().resolve()
        if p.exists():
            return p
        print(f"Path does not exist: {p}\nTry again.\n")


def infer_group_from_text(s: str) -> Optional[str]:
    m = re.search(r"group\s*0*([0-9]+)", s, flags=re.IGNORECASE)
    if m:
        return m.group(1).zfill(2)
    m = re.search(r"svg-?group\s*0*([0-9]+)", s, flags=re.IGNORECASE)
    if m:
        return m.group(1).zfill(2)
    return None


def ensure_exists(script_name: str):
    if not Path(script_name).exists():
        raise SystemExit(f"Missing script in current folder: {script_name}")


def find_group_json(work_dir: Path, group: str) -> Path:
    expected = work_dir / f"group{group}.json"
    if expected.exists():
        return expected
    # fallback: find any json in work_dir that contains group digits
    cands = sorted(work_dir.glob(f"*{group}*.json"))
    if cands:
        print(f"[WARN] Expected {expected.name} not found. Using: {cands[0].name}")
        return cands[0]
    raise SystemExit(f"Could not find extracted group JSON in {work_dir}")


def choose_join_keys(df_a: pd.DataFrame, df_b: pd.DataFrame) -> List[str]:
    """
    Picks best join keys that exist in BOTH dataframes.
    Prefer the most specific stable keys first.
    """
    candidates = [
        ["slide_id", "worker_id", "assignment_id"],
        ["slide_id", "assignment_id"],
        ["slide_id", "worker_id"],
        ["slide_id"],
    ]
    for keys in candidates:
        if all(k in df_a.columns for k in keys) and all(k in df_b.columns for k in keys):
            return keys

    # last-resort: intersection of common columns that look like IDs
    commons = [c for c in df_a.columns if c in df_b.columns]
    id_like = [c for c in commons if c.lower().endswith("_id") or c.lower() in ("worker_id", "assignment_id", "slide_id")]
    return id_like[:3]  # best effort


def merge_excels_by_sheet(
    preferences_xlsx: Path,
    intentionality_xlsx: Path,
    out_xlsx: Path,
    prefer_left: bool = True,
):
    """
    Merge columns into the SAME sheet:
    - For each sheet name in Preferences workbook, find same sheet in Intentionality workbook.
    - Merge on auto-chosen keys (usually slide_id, worker_id, assignment_id).
    - Avoid duplicate columns:
        * keep left columns as-is
        * add only columns from right that don't exist on left
        * if conflict, suffix right columns with '__intent'
    - Writes one merged workbook.
    Also copies over ERRORS sheets (if exist) as separate sheets.
    """

    pref_sheets = pd.read_excel(preferences_xlsx, sheet_name=None)
    intent_sheets = pd.read_excel(intentionality_xlsx, sheet_name=None)

    used_names = set()
    def safe_sheet_name(name: str) -> str:
        base = (name or "Sheet")[:31]
        if base not in used_names:
            used_names.add(base)
            return base
        i = 2
        while True:
            suffix = f"_{i}"
            cand = (base[:31 - len(suffix)] + suffix)[:31]
            if cand not in used_names:
                used_names.add(cand)
                return cand
            i += 1

    merge_report = []

    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as w:
        # merge normal sheets
        for sheet_name, df_pref in pref_sheets.items():
            # keep ERRORS separate
            if sheet_name.upper() == "ERRORS":
                continue

            df_int = intent_sheets.get(sheet_name)

            if df_int is None:
                # no matching sheet on the other side → write as-is
                out_name = safe_sheet_name(sheet_name)
                df_pref.to_excel(w, sheet_name=out_name, index=False)
                merge_report.append((sheet_name, "NO_MATCH_IN_INTENT", []))
                continue

            # choose keys
            keys = choose_join_keys(df_pref, df_int)
            if not keys:
                # can't merge safely → write both versions with suffix
                out_name1 = safe_sheet_name(sheet_name + "_PREF")
                out_name2 = safe_sheet_name(sheet_name + "_INTENT")
                df_pref.to_excel(w, sheet_name=out_name1, index=False)
                df_int.to_excel(w, sheet_name=out_name2, index=False)
                merge_report.append((sheet_name, "NO_KEYS_WROTE_BOTH", []))
                continue

            # prepare right-side columns (avoid duplicates)
            right_cols = [c for c in df_int.columns if c not in keys]
            df_right = df_int[keys + right_cols].copy()

            # rename right columns if conflict with left
            rename_map = {}
            for c in right_cols:
                if c in df_pref.columns:
                    rename_map[c] = f"{c}__intent"
            if rename_map:
                df_right = df_right.rename(columns=rename_map)

            merged = df_pref.merge(df_right, on=keys, how="left")

            out_name = safe_sheet_name(sheet_name)
            merged.to_excel(w, sheet_name=out_name, index=False)

            merge_report.append((sheet_name, "MERGED", keys))

        # copy ERRORS sheets if exist
        if "ERRORS" in pref_sheets:
            pref_sheets["ERRORS"].to_excel(w, sheet_name=safe_sheet_name("ERRORS_PREF"), index=False)
        if "ERRORS" in intent_sheets:
            intent_sheets["ERRORS"].to_excel(w, sheet_name=safe_sheet_name("ERRORS_INTENT"), index=False)

        # write merge report
        rep_df = pd.DataFrame(
            [{"sheet": s, "status": st, "join_keys": ", ".join(keys)} for (s, st, keys) in merge_report]
        )
        rep_df.to_excel(w, sheet_name=safe_sheet_name("MERGE_REPORT"), index=False)

    print(f"\n Merged workbook created: {out_xlsx}")
    print("   (See MERGE_REPORT sheet for join keys / status per sheet.)")


# ------------------ main ------------------

def main():
    # scripts you said exist
    for s in [
        "extract_json_by_group.py",
        "extract_json_by_slide.py",
        "pptx_shapes_to_svg_positions.py",
        "extract_motion_lines.py",
        "computeDistanceFromAllObj.py",
        "filter_dynamic_rows.py",
        "add_line_distances_to_intent.py",
        "add_static_distances_to_preferences.py",
        "build_intentionality_from_scratch.py",
    ]:
        ensure_exists(s)

    print("\n--- Full Pipeline Runner ---\n")
    big_json = ask_path("Enter path to the BIG JSON file: ")
    pptx_path = ask_path("Enter path to PPTX file OR a folder of PPTX files: ")

    inferred = infer_group_from_text(big_json.name) or infer_group_from_text(pptx_path.name)
    if inferred:
        group = inferred
        print(f"Inferred group = {group}")
    else:
        group = input("Enter group number (e.g. 08): ").strip().zfill(2)

    work_dir = Path(f"pipeline_output_group{group}").resolve()
    work_dir.mkdir(parents=True, exist_ok=True)

    slides_dir = work_dir / f"group{group}_slides_json"
    slides_dir.mkdir(exist_ok=True)
    

    shapes_out_dir = work_dir / "pptx_shapes_positions_csv"
    shapes_out_dir.mkdir(exist_ok=True)

    motion_lines_dir = work_dir / "motion_lines_metrics_csv"
    motion_lines_dir.mkdir(exist_ok=True)

    distances_dir = work_dir / "distances_csv"
    distances_dir.mkdir(exist_ok=True)

    # important outputs
    motion_intent_xlsx = work_dir / "motion_intent_analysis.xlsx"
    preferences_xlsx = work_dir / "Preferences_with_static_distances.xlsx"
    intentionality_xlsx = work_dir / "Intentionality.xlsx"
    final_merged_xlsx = work_dir / "FINAL_Merged_Preferences_Intentionality.xlsx"

    print(f"\nWork dir: {work_dir}\n")

    # 1) extract_json_by_group.py
    run_step(
        ["python3", "extract_json_by_group.py", group,
         "--input_file", str(big_json),
         "--output_dir", str(work_dir)],
        "extract_json_by_group.py"
    )

    group_json = find_group_json(work_dir, group)

    # 2) extract_json_by_slide.py
    run_step(
        ["python3", "extract_json_by_slide.py", group,
         "--input_file", str(group_json),
         "--output_dir", str(slides_dir)],
        "extract_json_by_slide.py"
    )

    inner = slides_dir / f"group{group}"
    if inner.exists():
        slides_input_for_dist = inner
    else:
        slides_input_for_dist = slides_dir

    # 3) pptx_shapes_to_svg_positions.py
    # supports --pptx OR --input-dir
    pptx_name_filter = None
    if pptx_path.is_file():
        run_step(
            ["python3", "pptx_shapes_to_svg_positions.py",
             "--pptx", str(pptx_path),
             "--output-dir", str(shapes_out_dir)],
            "pptx_shapes_to_svg_positions.py (single PPTX)"
        )
        pptx_name_filter = pptx_path.name
    else:
        run_step(
            ["python3", "pptx_shapes_to_svg_positions.py",
             "--input-dir", str(pptx_path),
             "--output-dir", str(shapes_out_dir)],
            "pptx_shapes_to_svg_positions.py (PPTX folder)"
        )

    # 4) extract_motion_lines.py
    run_step(
        ["python3", "extract_motion_lines.py",
         str(pptx_path),
         "--output-dir", str(motion_lines_dir)],
        "extract_motion_lines.py"
    )

    # 5) computeDistanceFromAllObj.py
    cmd = [
        "python3", "computeDistanceFromAllObj.py",
        str(slides_input_for_dist),
        "-o", str(distances_dir),
        "--centers-csv", str(shapes_out_dir),
        "--exclude-shapes", "Oval 8",
    ]
    if pptx_name_filter:
        cmd += ["--pptx-name", pptx_name_filter]
    run_step(cmd, "computeDistanceFromAllObj.py")

    # 6) filter_dynamic_rows.py  -> motion_intent_analysis.xlsx
    run_step(
        ["python3", "filter_dynamic_rows.py",
         str(distances_dir),
         "--output-file", str(motion_intent_xlsx)],
        "filter_dynamic_rows.py"
    )

    # 7) add_line_distances_to_intent.py
    # (writes back into the same file if your script supports overwriting via --output-file)
    line_csvs = list(motion_lines_dir.glob("*.csv"))

    if line_csvs:
        run_step(
            ["python3", "add_line_distances_to_intent.py",
             "--intent-file", str(motion_intent_xlsx),
             "--lines-file", str(motion_lines_dir),
             "--output-file", str(motion_intent_xlsx)],
            "add_line_distances_to_intent.py"
     )
    else:
        print("\n[WARN] No motion line CSV files found. Skipping add_line_distances_to_intent.py")

    # 8) add_static_distances_to_preferences.py  -> Preferences excel
    run_step(
        ["python3", "add_static_distances_to_preferences.py",
         "--input-xlsx", str(motion_intent_xlsx),
         "--csv-dir", str(distances_dir),
         "--output-xlsx", str(preferences_xlsx)],
        "add_static_distances_to_preferences.py"
    )

    # 9) build_intentionality_from_scratch.py -> Intentionality excel
    run_step(
        ["python3", "build_intentionality_from_scratch.py",
         "--input-xlsx", str(motion_intent_xlsx),
         "--csv-dir", str(distances_dir),
         "--output-xlsx", str(intentionality_xlsx)],
        "build_intentionality_from_scratch.py"
    )

    # 10) Merge Intentionality columns INTO Preferences (same sheets)
    final_merged_xlsx = work_dir / "RESULTS.xlsx"

    merge_intentionality_into_preferences(
        preferences_xlsx=preferences_xlsx,
        intentionality_xlsx=intentionality_xlsx,
        out_xlsx=final_merged_xlsx,
    )

    print(f"\n Final merged Excel: {final_merged_xlsx}")


if __name__ == "__main__":
    main()


#############
# HOW TO RUN:
# python3 run_full_pipeline.py
#
# Then enter:
#   - path to the big JSON file
#   - path to the PPTX file (or folder)
#
# FINAL OUTPUT:
#   pipeline_output_groupNN/RESULT.xlsx
#   (merged per sheet; see MERGE_REPORT inside)
#############
