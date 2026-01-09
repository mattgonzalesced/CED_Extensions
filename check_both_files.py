# -*- coding: utf-8 -*-
"""Check both files for duplicate IDs."""

import yaml
from collections import Counter

original_file = r"c:\CED_Extensions\AE PyDev.extension\AE pyTools.Tab\Test Buttons.Panel\Let there be JSON.pushbutton\Corporate_Full_Profile_mismatchremoved.yaml"
reordered_file = r"c:\CED_Extensions\AE PyDev.extension\AE pyTools.Tab\Test Buttons.Panel\Let there be JSON.pushbutton\Corporate_Full_Profile_mismatchremoved_REORDERED.yaml"

print("=" * 80)
print("CHECKING BOTH FILES FOR DUPLICATE IDs")
print("=" * 80)
print()

print("ORIGINAL FILE:")
with open(original_file, 'r', encoding='utf-8') as f:
    data1 = yaml.safe_load(f)
ids1 = [eq.get('id') for eq in data1['equipment_definitions']]
c1 = Counter(ids1)
print("  Total equipment definitions: {}".format(len(ids1)))
print("  Unique IDs: {}".format(len(set(ids1))))
dups1 = [(id, cnt) for id, cnt in c1.items() if cnt > 1]
if dups1:
    print("  DUPLICATES FOUND:")
    for id, cnt in dups1:
        print("    {} appears {} times".format(id, cnt))
else:
    print("  No duplicates")

print()
print("REORDERED FILE:")
with open(reordered_file, 'r', encoding='utf-8') as f:
    data2 = yaml.safe_load(f)
ids2 = [eq.get('id') for eq in data2['equipment_definitions']]
c2 = Counter(ids2)
print("  Total equipment definitions: {}".format(len(ids2)))
print("  Unique IDs: {}".format(len(set(ids2))))
dups2 = [(id, cnt) for id, cnt in c2.items() if cnt > 1]
if dups2:
    print("  DUPLICATES FOUND:")
    for id, cnt in dups2:
        print("    {} appears {} times".format(id, cnt))
else:
    print("  No duplicates")
