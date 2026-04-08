# -*- coding: utf-8 -*-
"""View-models for Add/Remove Spares and Spaces windows."""


class PanelListItem(object):
    """List row view-model for one panel schedule option."""

    def __init__(self, option, open_slots):
        self.option = option
        self.panel_id = int(option.get("panel_id", 0) or 0)
        self.panel_name = str(option.get("panel_name", "") or "Unnamed Panel")
        self.part_type = str(option.get("part_type_name", "") or option.get("board_type", "Unknown"))
        self.dist_system_name = str(option.get("dist_system_name", "") or "Unknown Dist. System")
        self.open_slots = int(max(0, open_slots or 0))
        self.open_slots_text = str(self.open_slots)
        self.action_text = ""
        self.is_checked = False


def action_label(action_type, mode):
    kind = str(action_type or "").lower()
    mode_key = str(mode or "").lower()
    if kind == "add":
        if mode_key == "spare":
            return "Add Spare"
        if mode_key == "space":
            return "Add Space"
        return "Add 50/50"
    if kind == "remove":
        if mode_key == "both":
            return "Remove Both"
        if mode_key == "space":
            return "Remove Space"
        return "Remove Spare"
    return ""
