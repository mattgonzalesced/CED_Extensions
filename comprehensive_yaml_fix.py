#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Comprehensive fix for corrupted YAML file."""

import re

input_file = r"c:\CED_Extensions\CEDLib.lib\prototypeHEB_StartCarrollton_Checkpoint35.yaml"
output_file = r"c:\CED_Extensions\CEDLib.lib\prototypeHEB_StartCarrollton_Checkpoint35_CLEANED.yaml"

with open(input_file, 'r', encoding='utf-8') as f:
    lines = f.readlines()

fixed_lines = []
skip_until_unindent = False
indent_level = 0

for i, line in enumerate(lines):
    # Check if line has malformed structures
    stripped = line.lstrip()
    current_indent = len(line) - len(stripped)

    # Skip lines with malformed YAML structures
    if re.match(r'^\s*\{\}:\s*$', line):
        skip_until_unindent = True
        indent_level = current_indent
        continue
    elif re.match(r'^\s*\[\]:\s*$', line):
        # Replace []: with empty list notation
        fixed_lines.append(line.replace('[]:', '[]'))
        continue
    elif re.match(r'^\s*\{\}\s*$', line):
        continue  # Skip standalone {}

    # If we're skipping nested structures, check if we've unindented
    if skip_until_unindent:
        if current_indent <= indent_level and stripped:
            skip_until_unindent = False
            indent_level = 0
            fixed_lines.append(line)
        # else continue skipping
    else:
        fixed_lines.append(line)

with open(output_file, 'w', encoding='utf-8') as f:
    f.writelines(fixed_lines)

print("Cleaned YAML file")
print("Original lines: {}".format(len(lines)))
print("Cleaned lines:  {}".format(len(fixed_lines)))
print("Removed lines:  {}".format(len(lines) - len(fixed_lines)))
print("Output written to: {}".format(output_file))
