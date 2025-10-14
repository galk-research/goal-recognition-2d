#!/usr/bin/env python3\
import json
import os
import argparse

 """
    Splits a JSON file for a specific group into individual files per slide.

    Parameters:
    - input_file (str): Path to the input JSON file for the group.
    - output_dir (str): Directory to save the extracted slide files.
    - group_number (int): The group number to determine the slide labels.
"""

def split_json_by_slide(input_file, output_dir, group_number):
    try:
        start_label = f"svg-group{group_number:02}_slide01"

        # Read the JSON file
        with open(input_file, 'r') as infile:
            data = json.load(infile)

        if not isinstance(data, dict):
            raise ValueError("The JSON structure must be a dictionary for this operation.")

        # Create a new directory for this group
        group_dir = os.path.join(output_dir, f"group{group_number:02}")
        os.makedirs(group_dir, exist_ok=True)

        # Iterate over the dictionary's keys to separate data by slide
        for key in data.keys():
            if key.startswith(f"svg-group{group_number:02}_slide"):
                slide_number = key.split("_")[-1]
                slide_data = {key: data[key]}
                output_file = os.path.join(group_dir, f"{key}.json")
                with open(output_file, 'w') as outfile:
                    json.dump(slide_data, outfile, indent=4)

                print(f"Saved {key} to {output_file}")

    except Exception as e:
        print(f"An error occurred: {e}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Split JSON data by slide number.")
    parser.add_argument("group_number", type=int, help="The group number to split.")
    parser.add_argument(
        "--input_file",
        default=" ",
        help="Path to the input JSON file (default: specified path in code)."
    )
    parser.add_argument(
        "--output_dir",
        default=" ",  # Set a default directory
        help="Directory to save the output files (default: specified path in code)."
    )
    args = parser.parse_args()

    split_json_by_slide(args.input_file, args.output_dir, args.group_number)



################
# HOW TO RUN:
# run - python3 name_of_file.py <group_number> [--input_file <input_file_path>] [--output_dir <output_directory>]
# Explanation of Arguments:
# <group_number> (Required):
# This is the group number you want to split into individual files per slide


# --input_file <input_file_path> (Optional):
# If you want to use a different input JSON file than the default, specify the file path here.
# Example: --input_file /path/to/your/input_file.json
# If you don't provide this argument, the default path that you provide above will be used.

# --output_dir <output_directory> (Optional):
# This is where the output file will be saved.
# Example: --output_dir /path/to/save/directory/
# If you don't specify this, the output will be saved to the default directory that you provide above.

