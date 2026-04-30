# -*- coding: utf-8 -*-
"""
One-off converter: HEB_profiles_V4_MODIFIED_41.yaml -> HEB_profiles_V5.yaml.

Schema change v4 -> v100:
    * Each LED's three peer lists (tags / keynotes / text_notes) are
      collapsed into a single ``annotations`` list. Each entry carries
      a ``kind`` field ("tag" / "keynote" / "text_note"), a generated
      ``id`` of the form ``{led_id}-ANN-{NNN}``, and a ``label`` derived
      from family/type (or the trimmed text content for text_notes).
    * For text_notes, the original ``text`` field is preserved at the
      top level — Place Element Annotations needs it to recreate the
      note on placement.
    * schema_version is stamped to 100 at the document root and on
      every equipment_definitions entry.
    * The legacy peer lists are dropped.

All other fields (parameters, offsets, parent_filter,
equipment_properties, etc.) are preserved verbatim.
"""

from __future__ import print_function

import io
import os
import sys

import yaml


SRC = r"c:\CED_Extensions\HEB_profiles_V4_MODIFIED_41.yaml"
DST = r"c:\CED_Extensions\HEB_profiles_V5.yaml"


def _label_for_text(text):
    text = (text or "").strip()
    if not text:
        return "(empty text note)"
    return (text[:60] + "...") if len(text) > 60 else text


def _label_for_typed(entry, kind):
    family = (entry.get("family_name") or "").strip()
    type_name = (entry.get("type_name") or "").strip()
    if family and type_name and family.lower() != "null" and type_name.lower() != "null":
        return "{} : {}".format(family, type_name)
    cat = (entry.get("category_name") or "").strip()
    if cat and cat.lower() != "null":
        return cat
    return kind


def _make_annotation(entry, kind, ann_id):
    """Convert one legacy tag/keynote/text_note entry to a v100 annotation."""
    out = {
        "id": ann_id,
        "kind": kind,
    }

    # Carry over everything that already exists, with stable ordering.
    for key in ("category_name", "family_name", "type_name", "type",
                "parameters", "offsets", "leaders", "width_inches"):
        if key in entry:
            out[key] = entry[key]

    if kind == "text_note":
        text_content = (entry.get("text") or "").strip()
        out["text"] = text_content
        out["label"] = _label_for_text(text_content)
    else:
        out["label"] = _label_for_typed(entry, kind)

    return out


def _convert_led(led, led_id):
    """Mutate ``led`` in place: build annotations list, drop peer lists,
    leave everything else untouched."""
    if not isinstance(led, dict):
        return 0

    annotations = []
    counter = 0

    def _next_id():
        return "{}-ANN-{:03d}".format(led_id, counter + 1)

    for kind, key in (("tag", "tags"),
                      ("keynote", "keynotes"),
                      ("text_note", "text_notes")):
        peer = led.get(key)
        if not isinstance(peer, list):
            continue
        for entry in peer:
            if not isinstance(entry, dict):
                continue
            annotations.append(_make_annotation(entry, kind, _next_id()))
            counter += 1

    led["annotations"] = annotations
    for key in ("tags", "keynotes", "text_notes"):
        led.pop(key, None)
    return counter


def _convert_profile(profile):
    if not isinstance(profile, dict):
        return 0
    profile["schema_version"] = 100
    converted = 0
    for set_dict in profile.get("linked_sets") or []:
        if not isinstance(set_dict, dict):
            continue
        for led in set_dict.get("linked_element_definitions") or []:
            if not isinstance(led, dict):
                continue
            led_id = led.get("id") or ""
            converted += _convert_led(led, led_id)
    return converted


def main():
    print("Reading {} ...".format(SRC))
    with io.open(SRC, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    if not isinstance(data, dict):
        print("ERROR: top-level YAML is not a mapping", file=sys.stderr)
        return 1

    data["schema_version"] = 100
    profiles = data.get("equipment_definitions") or []
    print("Found {} profile(s).".format(len(profiles)))

    total_anns = 0
    for profile in profiles:
        total_anns += _convert_profile(profile)

    print("Built {} annotation entries across all LEDs.".format(total_anns))

    print("Writing {} ...".format(DST))
    with io.open(DST, "w", encoding="utf-8") as f:
        yaml.safe_dump(
            data,
            f,
            default_flow_style=False,
            sort_keys=False,
            allow_unicode=True,
            width=4096,
        )

    print("Done. Bytes written: {}".format(os.path.getsize(DST)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
