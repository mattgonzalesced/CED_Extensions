#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Fix malformed nested empty dictionaries in YAML file."""

import re

input_file = r"c:\CED_Extensions\CEDLib.lib\prototypeHEB_StartCarrollton_Checkpoint35.yaml"
output_file = r"c:\CED_Extensions\CEDLib.lib\prototypeHEB_StartCarrollton_Checkpoint35_FIXED.yaml"

with open(input_file, 'r', encoding='utf-8') as f:
    content = f.read()

# Pattern to match any field with nested {} structures
# Matches "field_name:" followed by nested {}:
pattern1 = r'(\s+\w+:)\s*\n(\s+\{\}:\s*\n)+'
replacement1 = r'\1 {}\n'
fixed_content = re.sub(pattern1, replacement1, content)

# Also remove any standalone lines with just whitespace and {}:
pattern2 = r'^\s+\{\}:\s*$'
fixed_content = re.sub(pattern2, '', fixed_content, flags=re.MULTILINE)

# Remove standalone lines with just {}
pattern3 = r'^\s+\{\}\s*$'
fixed_content = re.sub(pattern3, '', fixed_content, flags=re.MULTILINE)

with open(output_file, 'w', encoding='utf-8') as f:
    f.write(fixed_content)

print("Fixed {} fields with nested {{}} structures".format(
    len(re.findall(pattern1, content))
))
print("Removed {} standalone {{}}:lines".format(
    len(re.findall(pattern2, content, flags=re.MULTILINE))
))
print("Removed {} standalone {{}} lines".format(
    len(re.findall(pattern3, content, flags=re.MULTILINE))
))
print("Output written to: {}".format(output_file))
