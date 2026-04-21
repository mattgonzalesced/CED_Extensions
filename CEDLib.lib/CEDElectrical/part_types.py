# -*- coding: utf-8 -*-
"""Shared family part-type constants for CEDElectrical modules."""

# Revit BuiltInParameter.FAMILY_CONTENT_PART_TYPE values
PART_TYPE_PANELBOARD = 14
PART_TYPE_TRANSFORMER = 15
PART_TYPE_SWITCHBOARD = 16
PART_TYPE_OTHER_PANEL = 17
PART_TYPE_EQUIPMENT_SWITCH = 18

PART_TYPE_MAP = {
    PART_TYPE_PANELBOARD: "Panelboard",
    PART_TYPE_TRANSFORMER: "Transformer",
    PART_TYPE_SWITCHBOARD: "Switchboard",
    PART_TYPE_OTHER_PANEL: "Other Panel",
    PART_TYPE_EQUIPMENT_SWITCH: "Equipment Switch",
}


def part_type_name(part_type, default_name="Unknown"):
    """Return display label for family part type integer."""
    return PART_TYPE_MAP.get(part_type, default_name)

