# -*- coding: utf-8 -*-
"""
Normalize Truth Groups
----------------------
One-shot normalization for active YAML truth-group metadata and payload sync.
"""

import os
import sys

from pyrevit import forms, script
output = script.get_output()
output.close_others()

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from ExtensibleStorage.yaml_store import normalize_active_yaml_truth_groups  # noqa: E402
from LogicClasses.yaml_path_cache import get_yaml_display_name  # noqa: E402

TITLE = "Normalize Truth Groups"


def main():
    try:
        yaml_path, report, saved = normalize_active_yaml_truth_groups(None, action=TITLE)
    except RuntimeError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    except Exception as exc:
        forms.alert("Failed to normalize truth groups:\n\n{}".format(exc), title=TITLE)
        return

    yaml_label = get_yaml_display_name(yaml_path)
    lines = [
        "YAML source: {}".format(yaml_label),
        "Changed: {}".format("Yes" if saved else "No"),
        "",
        "Profiles scanned: {}".format(report.get("total_profiles", 0)),
        "Groups scanned: {}".format(report.get("total_groups", 0)),
        "Groups with drift detected: {}".format(report.get("groups_with_drift", 0)),
        "Groups promoted from member edits: {}".format(report.get("groups_promoted_from_member", 0)),
        "Groups with conflicting member edits: {}".format(
            report.get("groups_with_conflicting_member_changes", 0)
        ),
        "Profiles payload synced: {}".format(report.get("profiles_payload_synced", 0)),
        "Profiles metadata repaired: {}".format(report.get("profiles_metadata_repaired", 0)),
    ]
    if saved:
        lines.append("")
        lines.append("Normalized truth-group data has been saved back to Extensible Storage.")
    else:
        lines.append("")
        lines.append("No changes were required.")
    forms.alert("\n".join(lines), title=TITLE)


if __name__ == "__main__":
    main()
