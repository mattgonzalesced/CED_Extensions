# -*- coding: utf-8 -*-
"""Graphics and category-visibility helpers for Circuit Element Finder."""

from pyrevit import DB


HIDE_MEP_CATEGORY_IDS = [
    DB.BuiltInCategory.OST_DuctCurves,
    DB.BuiltInCategory.OST_DuctFitting,
    DB.BuiltInCategory.OST_DuctTerminal,
    DB.BuiltInCategory.OST_PipeCurves,
    DB.BuiltInCategory.OST_PipeFitting,
    DB.BuiltInCategory.OST_PipeAccessory,
    DB.BuiltInCategory.OST_MechanicalEquipment,
]


def hide_mep_categories(view, doc, logger=None, include_mechanical_equipment=True):
    """Hide duct/pipe categories in the target view."""
    hidden_count = 0
    for bic in list(HIDE_MEP_CATEGORY_IDS or []):
        if (not include_mechanical_equipment) and bic == DB.BuiltInCategory.OST_MechanicalEquipment:
            continue
        category = None
        try:
            category = DB.Category.GetCategory(doc, bic)
        except Exception:
            category = None
        if category is None:
            continue
        try:
            if not view.CanCategoryBeHidden(category.Id):
                continue
        except Exception:
            continue
        try:
            already_hidden = bool(view.GetCategoryHidden(category.Id))
        except Exception:
            already_hidden = False
        if already_hidden:
            continue
        try:
            view.SetCategoryHidden(category.Id, True)
            hidden_count += 1
        except Exception as ex:
            if logger:
                logger.debug("Category hide failed for {0}: {1}".format(category.Name, ex))
    return hidden_count


def build_selection_override_settings(line_color=None, line_weight=8):
    """Build strong projection override settings."""
    color = line_color
    if color is None:
        color = DB.Color(240, 40, 40)
    override = DB.OverrideGraphicSettings()
    try:
        override.SetProjectionLineColor(color)
    except Exception:
        pass
    try:
        override.SetProjectionLineWeight(int(line_weight or 8))
    except Exception:
        pass
    try:
        override.SetHalftone(False)
    except Exception:
        pass
    try:
        override.SetSurfaceTransparency(0)
    except Exception:
        pass
    return override


def apply_selection_overrides(view, element_ids, line_color=None, line_weight=8, logger=None):
    """Apply element-level overrides to emphasize the selected elements."""
    override = build_selection_override_settings(line_color=line_color, line_weight=line_weight)
    applied = 0
    for element_id in list(element_ids or []):
        try:
            view.SetElementOverrides(element_id, override)
            applied += 1
        except Exception as ex:
            if logger:
                logger.debug("Override failed for element {0}: {1}".format(element_id, ex))
    return applied
