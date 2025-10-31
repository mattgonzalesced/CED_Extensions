# -*- coding: utf-8 -*-
__title__ = 'Update Custom\nDictionary'
__doc__ = 'Copy custom dictionary file to Revit 2024 spelling dictionary location'

import os
import shutil
from pyrevit import script

# Get the directory where this script is located
script_dir = os.path.dirname(__file__)
source_file = os.path.join(script_dir, 'Custom.dic')

# Target locations for Revit years 2017-2026
years = range(2017, 2027)
target_paths = [r"C:\ProgramData\Autodesk\RVT {}\Custom.dic".format(year) for year in years]

# Get logger
output = script.get_output()

# Check if source file exists
if not os.path.exists(source_file):
    output.print_md("**Error:** Source file not found: {}".format(source_file))
else:
    success_count = 0
    skipped_count = 0
    errors = []

    for target_path in target_paths:
        try:
            target_dir = os.path.dirname(target_path)

            # Skip if directory doesn't exist
            if not os.path.exists(target_dir):
                skipped_count += 1
                continue

            # Copy the file
            shutil.copy2(source_file, target_path)
            success_count += 1

        except Exception as e:
            errors.append("Error copying to {}: {}".format(target_path, str(e)))

    # Print results
    output.print_md("**Custom Dictionary Update Complete**\n")
    output.print_md("- Successfully updated: **{}** locations".format(success_count))
    output.print_md("- Skipped (folder not found): **{}** locations".format(skipped_count))

    if errors:
        output.print_md("\n**Errors:**")
        for error in errors:
            output.print_md("- {}".format(error))
