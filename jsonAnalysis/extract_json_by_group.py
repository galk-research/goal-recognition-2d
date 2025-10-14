#!/usr/bin/env python3\
import json
import os
import argparse

"""
    Extracts a section of a JSON file based on dynamic group labels.

    Parameters:
    - input_file (str): Path to the input JSON file.
    - output_file (str): Path to save the extracted JSON section.
    - group_number (int): The group number to determine the labels.
"""

def extract_json_by_group(input_file, output_file, group_number):
    try:
        start_label = f"svg-group{group_number:02}_slide01"
        stop_label = f"svg-group{group_number + 1:02}_slide01"

        with open(input_file, 'r') as infile:
            data = json.load(infile)
   
        if not isinstance(data, dict):
            raise ValueError("The JSON structure must be a dictionary for this operation.")

        # Extract data between start_label and stop_label
        extracting = False
        extracted_data = {}

        # Iterate over the dictionary's keys in order
        for key in data.keys():
            if key == stop_label: 
                break

            if key == start_label:
                extracting = True
            
            if extracting:
                extracted_data[key] = data[key]

        os.makedirs(os.path.dirname(output_file), exist_ok=True)

        # Write the extracted data to the output file
        with open(output_file, 'w') as outfile:
            json.dump(extracted_data, outfile, indent=4)

        print(f"Extracted data from {start_label} to {stop_label} saved to {output_file}.")

    except Exception as e:
         print(f"An error occurred: {e}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract JSON data by group number.")
    parser.add_argument("group_number", type=int, help="The group number to extract.")
    parser.add_argument(
        "--input_file",
        default=" ", ## Set a default path to the JSON file
        help="Path to the input JSON file (default: specified path in code)."
    )
    parser.add_argument(
        "--output_dir",
        default=" ",  # Set a default directory
        help="Directory to save the output JSON files (default: specified path in code)."
    )
    args = parser.parse_args()

   # Dynamically create the output file name based on the group number
    output_file = os.path.join(args.output_dir, f"group{args.group_number:02}_from_json.json")

    extract_json_by_group(args.input_file, output_file, args.group_number)       


################
# HOW TO RUN:
# run - python3 name_of_file.py <group_number> [--input_file <input_file_path>] [--output_dir <output_directory>]
# Explanation of Arguments:
# <group_number> (Required):
# This is the group number you want to extract.


# --input_file <input_file_path> (Optional):
# If you want to use a different input JSON file than the default, specify the file path here.
# Example: --input_file /path/to/your/input_file.json
# If you don't provide this argument, the default path that you provide above will be used.

# --output_dir <output_directory> (Optional):
# This is where the output file will be saved.
# Example: --output_dir /path/to/save/directory/
# If you don't specify this, the output will be saved to the default directory that you provide above.


