# -*- coding: utf-8 -*-
"""Compare two YAML files to check for duplicates and differences."""

import yaml

original_file = r"c:\CED_Extensions\AE PyDev.extension\AE pyTools.Tab\Test Buttons.Panel\Let there be JSON.pushbutton\Corporate_Full_Profile_mismatchremoved.yaml"
reordered_file = r"c:\CED_Extensions\AE PyDev.extension\AE pyTools.Tab\Test Buttons.Panel\Let there be JSON.pushbutton\Corporate_Full_Profile_mismatchremoved_REORDERED.yaml"

with open(original_file, 'r', encoding='utf-8') as f:
    data1 = yaml.safe_load(f)

with open(reordered_file, 'r', encoding='utf-8') as f:
    data2 = yaml.safe_load(f)

eq1 = data1['equipment_definitions']
eq2 = data2['equipment_definitions']

print("=" * 80)
print("YAML COMPARISON")
print("=" * 80)
print()
print("Original file:  {} equipment definitions".format(len(eq1)))
print("Reordered file: {} equipment definitions".format(len(eq2)))
print()

# Find "PF_Plan TV Video Wall" entries
tv_orig = [eq for eq in eq1 if 'PF_Plan TV Video Wall' in eq.get('name', '')]
tv_reord = [eq for eq in eq2 if 'PF_Plan TV Video Wall' in eq.get('name', '')]

print("PF_Plan TV Video Wall entries:")
print("  Original:  {}".format(len(tv_orig)))
print("  Reordered: {}".format(len(tv_reord)))
print()

print("ORIGINAL FILE entries with 'PF_Plan TV Video Wall':")
for eq in tv_orig:
    led_count = 0
    if eq.get('linked_sets'):
        led_count = len(eq.get('linked_sets', [{}])[0].get('linked_element_definitions', []))
    print("  ID: {}, Name: {}, LEDs: {}, Category: {}".format(
        eq.get('id'),
        eq.get('name'),
        led_count,
        eq.get('parent_filter', {}).get('category')
    ))

print()
print("REORDERED FILE entries with 'PF_Plan TV Video Wall':")
for eq in tv_reord:
    led_count = 0
    if eq.get('linked_sets'):
        led_count = len(eq.get('linked_sets', [{}])[0].get('linked_element_definitions', []))
    print("  ID: {}, Name: {}, LEDs: {}, Category: {}".format(
        eq.get('id'),
        eq.get('name'),
        led_count,
        eq.get('parent_filter', {}).get('category')
    ))

print()
print("=" * 80)

# Check all IDs
ids1 = [eq.get('id') for eq in eq1]
ids2 = [eq.get('id') for eq in eq2]

print("All IDs in original:")
for i, eq_id in enumerate(ids1, 1):
    print("  {}: {}".format(i, eq_id))

print()
print("All IDs in reordered:")
for i, eq_id in enumerate(ids2, 1):
    print("  {}: {}".format(i, eq_id))
