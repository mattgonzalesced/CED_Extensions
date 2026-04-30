# -*- coding: utf-8 -*-
"""
One-off patch: fix bogus ``is_group: true`` flags in HEB_profiles_V5.yaml.

The legacy v4 source set ``is_group: true`` on ~99% of LEDs regardless of
whether they were actually Revit Groups. The new placement engine takes
the flag literally and tries Doc.Create.PlaceGroup, which fails for
LEDs whose label is in ``Family : Type`` format (a FamilyInstance
marker). Heuristic: if the label contains `` : ``, the LED is a family
and ``is_group`` must be false.
"""

from __future__ import print_function

import io
import sys

import yaml


PATH = r"c:\CED_Extensions\HEB_profiles_V5.yaml"
LABEL_SEP = " : "


def main():
    print("Reading {} ...".format(PATH))
    with io.open(PATH, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    flipped = 0
    left_true = 0
    total_leds = 0
    for profile in data.get("equipment_definitions") or []:
        if not isinstance(profile, dict):
            continue
        for set_dict in profile.get("linked_sets") or []:
            if not isinstance(set_dict, dict):
                continue
            for led in set_dict.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                total_leds += 1
                label = led.get("label") or ""
                if led.get("is_group") and LABEL_SEP in label:
                    led["is_group"] = False
                    flipped += 1
                elif led.get("is_group"):
                    left_true += 1

    print("LEDs scanned:        {}".format(total_leds))
    print("Flipped to false:    {}".format(flipped))
    print("Left as is_group=true: {}  (label has no ' : ' — likely real groups)".format(left_true))

    print("Writing patched file back to {} ...".format(PATH))
    with io.open(PATH, "w", encoding="utf-8") as f:
        yaml.safe_dump(
            data,
            f,
            default_flow_style=False,
            sort_keys=False,
            allow_unicode=True,
            width=4096,
        )
    print("Done.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
