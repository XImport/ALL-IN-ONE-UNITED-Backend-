import os
from pathlib import Path
from datetime import datetime


def change_dict_key(original_dict, key_to_change, new_key):
    new_dict = {}
    for key, value in original_dict.items():
        if key == key_to_change:
            new_dict[new_key] = value
        else:
            new_dict[key] = value

    return new_dict


def OrderExcelFilesByNames(directory):
    # Get all files in the directory
    filenames = [file for file in os.listdir(directory) if file.endswith(".xlsx")]

    # Extract the month and year from each filename and store as (filename, month, year) tuple
    file_months_years = []
    for filename in filenames:
        print(filename)
        try:
            month, year = map(int, filename[:-5].split("-"))
            file_months_years.append((filename, month, year))
        except (ValueError, IndexError):
            print(f"Skipping invalid filename: {filename}")

    # Sort the list of (filename, month, year) tuples based on month and year
    sorted_file_months_years = sorted(file_months_years, key=lambda x: (x[2], x[1]))

    # Extract the sorted filenames
    sorted_filenames = [pair[0] for pair in sorted_file_months_years]

    return sorted_filenames
