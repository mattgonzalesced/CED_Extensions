# -*- coding: utf-8 -*-
"""Restore parent directive relationships in HEB_profiles_V5.2.yaml.

Some V4 -> V5 migrations stringified parent directives so values that
should look like::

    FLA Input_CED:
      parent_parameter: CED-E-AMPS

ended up as opaque strings::

    FLA Input_CED: 'parent_parameter: \\"CED-E-AMPS\\"'

That string passes through ``_apply_static_parameters`` as a literal
text value at placement time — i.e. it overwrites the receptacle's
amperage parameter with a useless quoted string instead of routing
through ``directives.parent_directive`` at audit time.

This script walks every LED's ``parameters`` dict, detects values
matching the legacy ``parent_parameter: "<name>"`` shape (with or
without escaped quotes), and rewrites them as the canonical V5
directive dict. Sibling-directive strings (``sibling_parameter: ...``)
get the same treatment for completeness.

Output goes to ``HEB_profiles_V5.2.yaml`` (in place); a sibling
``.bak`` copy is written first so the original is recoverable.
"""

import io
import os
import re
import shutil

import yaml


SCRIPT_DIR = r"c:\CED_Extensions"
IN_PATH = os.path.join(SCRIPT_DIR, "HEB_profiles_V5.2.yaml")
BACKUP_PATH = IN_PATH + ".bak"


# Match ``parent_parameter: "X"`` or ``parent_parameter: \"X\"`` or
# unquoted (``parent_parameter: X``). Whitespace tolerant. The capture
# group is the parameter name only — without surrounding quotes /
# backslashes.
_PARENT_RE = re.compile(
    r"""^\s*parent_parameter\s*:\s*\\?["']?(?P<name>[^"'\\]+?)\\?["']?\s*$"""
)
_SIBLING_RE = re.compile(
    r"""^\s*sibling_parameter\s*:\s*\\?["']?(?P<name>[^"'\\]+?)\\?["']?\s*$"""
)


def _coerce_directive(value):
    """If ``value`` is a legacy stringified directive, return the dict
    form. Otherwise return ``value`` unchanged."""
    if not isinstance(value, str):
        return value
    m = _PARENT_RE.match(value)
    if m:
        return {"parent_parameter": m.group("name").strip()}
    m = _SIBLING_RE.match(value)
    if m:
        return {"sibling_parameter": m.group("name").strip()}
    return value


def _walk_and_fix(profile_data):
    """Walk every LED's ``parameters`` dict and rewrite legacy
    stringified directives in place. Returns the number of values
    converted."""
    converted = 0
    for profile in (profile_data or {}).get("equipment_definitions") or []:
        if not isinstance(profile, dict):
            continue
        for s in profile.get("linked_sets") or []:
            if not isinstance(s, dict):
                continue
            for led in s.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                params = led.get("parameters")
                if not isinstance(params, dict):
                    continue
                for name, value in list(params.items()):
                    new_value = _coerce_directive(value)
                    if new_value is not value:
                        params[name] = new_value
                        converted += 1
    return converted


def main():
    with io.open(IN_PATH, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    if not isinstance(data, dict):
        print("ERROR: V5.2 YAML didn't parse to a mapping.")
        return 1

    converted = _walk_and_fix(data)
    print("Converted {} stringified directive(s) to dict form.".format(converted))
    if converted == 0:
        print("Nothing to do.")
        return 0

    # Backup first.
    if not os.path.exists(BACKUP_PATH):
        shutil.copyfile(IN_PATH, BACKUP_PATH)
        print("Backed up original to {}".format(BACKUP_PATH))
    else:
        print("Backup already exists at {} (not overwritten).".format(BACKUP_PATH))

    with io.open(IN_PATH, "w", encoding="utf-8") as f:
        yaml.safe_dump(
            data, f,
            default_flow_style=False,
            sort_keys=False,
            allow_unicode=True,
            width=10**9,
        )
    size = os.path.getsize(IN_PATH)
    print("Wrote {} ({} bytes).".format(IN_PATH, size))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
