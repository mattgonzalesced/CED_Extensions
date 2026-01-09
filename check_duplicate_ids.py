# -*- coding: utf-8 -*-
"""Check for duplicate IDs in YAML file."""

import yaml
from collections import Counter

original_file = r"c:\CED_Extensions\AE PyDev.extension\AE pyTools.Tab\Test Buttons.Panel\Let there be JSON.pushbutton\Corporate_Full_Profile_mismatchremoved.yaml"

with open(original_file, 'r', encoding='utf-8') as f:
    data = yaml.safe_load(f)

eq_defs = data['equipment_definitions']
ids = [eq.get('id') for eq in eq_defs]
counts = Counter(ids)

print("=" * 80)
print("CHECKING FOR DUPLICATE IDs")
print("=" * 80)
print()
print("Total equipment definitions: {}".format(len(ids)))
print("Unique IDs: {}".format(len(set(ids))))
print()

duplicates = [(eq_id, count) for eq_id, count in counts.items() if count > 1]

if duplicates:
    print("DUPLICATE IDs FOUND:")
    for eq_id, count in sorted(duplicates):
        print("  {} appears {} times".format(eq_id, count))

    print()
    print("Details of duplicate entries:")
    for eq_id, count in sorted(duplicates):
        print()
        print("  ID: {}".format(eq_id))
        matching = [eq for eq in eq_defs if eq.get('id') == eq_id]
        for i, eq in enumerate(matching, 1):
            led_count = 0
            if eq.get('linked_sets'):
                led_count = len(eq.get('linked_sets', [{}])[0].get('linked_element_definitions', []))
            print("    Entry {}: Name='{}', Category='{}', LEDs={}".format(
                i,
                eq.get('name'),
                eq.get('parent_filter', {}).get('category'),
                led_count
            ))
else:
    print("No duplicate IDs found - all equipment definitions have unique IDs")
