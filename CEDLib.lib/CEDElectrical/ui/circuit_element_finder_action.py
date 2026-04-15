# -*- coding: utf-8 -*-
"""Circuit Element Finder action orchestration."""

import Autodesk.Revit.DB.Electrical as DBE
from pyrevit import DB, forms, revit, script

from CEDElectrical.Application.services.circuit_element_finder_graphics import (
    apply_selection_overrides,
    hide_mep_categories,
)
from CEDElectrical.Application.services.circuit_element_finder_view_finder import (
    find_existing_3d_view,
    find_existing_plan_view,
    get_or_create_dedicated_view,
)
from CEDElectrical.Domain.circuit_element_finder_bounds import (
    expand_bounding_box,
    get_combined_bounding_box,
    show_elements,
)
from Snippets import _elecutils as eu
from Snippets.circuit_ui_actions import collect_circuit_targets, set_revit_selection

OPTION_EXISTING_PLAN = "Open in Existing Plan View"
OPTION_EXISTING_3D = "Open in Existing 3D View"
OPTION_DEDICATED_PLAN = "Create / Use Dedicated Plan View"
OPTION_DEDICATED_3D = "Create / Use Dedicated 3D View"
OPTION_SHOW_ELEMENTS_ONLY = "Show Elements Only (No View Selection)"

VALID_SELECTION_BICS = (
    DB.BuiltInCategory.OST_ElectricalEquipment,
    DB.BuiltInCategory.OST_ElectricalFixtures,
    DB.BuiltInCategory.OST_LightingDevices,
    DB.BuiltInCategory.OST_LightingFixtures,
)


def _id_value(item):
    try:
        return int(item.IntegerValue)
    except Exception:
        return -1


def _bic_id_value(bic):
    try:
        return int(DB.ElementId(bic).IntegerValue)
    except Exception:
        try:
            return int(bic)
        except Exception:
            return 0


def _is_valid_level_id(doc, level_id):
    if not isinstance(level_id, DB.ElementId):
        return False
    if level_id == DB.ElementId.InvalidElementId:
        return False
    try:
        level = doc.GetElement(level_id)
    except Exception:
        level = None
    return isinstance(level, DB.Level)


def _element_level_id(doc, element):
    if element is None:
        return DB.ElementId.InvalidElementId

    level_id = getattr(element, "LevelId", None)
    if _is_valid_level_id(doc, level_id):
        return level_id

    level_bips = []
    for name in ("INSTANCE_REFERENCE_LEVEL_PARAM", "FAMILY_LEVEL_PARAM", "RBS_START_LEVEL_PARAM"):
        try:
            bip = getattr(DB.BuiltInParameter, name)
        except Exception:
            bip = None
        if bip is not None:
            level_bips.append(bip)
    for bip in list(level_bips):
        try:
            param = element.get_Parameter(bip)
        except Exception:
            param = None
        if not param:
            continue
        try:
            maybe_id = param.AsElementId()
        except Exception:
            maybe_id = None
        if _is_valid_level_id(doc, maybe_id):
            return maybe_id

    return DB.ElementId.InvalidElementId


def _collect_valid_selection(doc, uidoc):
    allowed_category_ids = set([_bic_id_value(x) for x in VALID_SELECTION_BICS])

    valid_ids = []
    valid_elements = []
    category_ids = []
    seen = set()
    seen_category_ids = set()
    try:
        selected_ids = list(uidoc.Selection.GetElementIds() or [])
    except Exception:
        selected_ids = []
    for element_id in selected_ids:
        if not isinstance(element_id, DB.ElementId):
            continue
        eid_val = _id_value(element_id)
        if eid_val <= 0 or eid_val in seen:
            continue
        seen.add(eid_val)

        try:
            element = doc.GetElement(element_id)
        except Exception:
            element = None
        if element is None:
            continue
        category = getattr(element, "Category", None)
        category_id = _id_value(getattr(category, "Id", None))
        if category_id not in allowed_category_ids:
            continue

        valid_ids.append(element_id)
        valid_elements.append(element)
        try:
            cat_id_obj = getattr(category, "Id", None)
            cat_id_val = _id_value(cat_id_obj)
        except Exception:
            cat_id_obj = None
            cat_id_val = -1
        if isinstance(cat_id_obj, DB.ElementId) and cat_id_val > 0 and cat_id_val not in seen_category_ids:
            seen_category_ids.add(cat_id_val)
            category_ids.append(cat_id_obj)

    return valid_ids, valid_elements, category_ids


def _preferred_level_id(doc, elements):
    for element in list(elements or []):
        maybe_level_id = _element_level_id(doc, element)
        if _is_valid_level_id(doc, maybe_level_id):
            return maybe_level_id
    return DB.ElementId.InvalidElementId


def _activate_view(uidoc, view, logger=None):
    if view is None:
        return False
    try:
        if uidoc.ActiveView is not None and uidoc.ActiveView.Id == view.Id:
            return True
    except Exception:
        pass

    try:
        uidoc.ActiveView = view
    except Exception as ex:
        if logger:
            logger.debug("ActiveView set failed: {0}".format(ex))
        return False
    try:
        return bool(uidoc.ActiveView is not None and uidoc.ActiveView.Id == view.Id)
    except Exception:
        return False


def _resolve_target_view(
    doc,
    mode,
    preferred_level_id,
    required_category_ids=None,
    selected_element_ids=None,
    require_all_visible=True,
    logger=None,
):
    selected = str(mode or "").strip()
    if selected == OPTION_EXISTING_PLAN:
        view = find_existing_plan_view(
            doc,
            preferred_level_id=preferred_level_id,
            required_category_ids=required_category_ids,
            selected_element_ids=selected_element_ids,
            require_all_visible=require_all_visible,
        )
        return view, False, False
    if selected == OPTION_EXISTING_3D:
        view = find_existing_3d_view(
            doc,
            required_category_ids=required_category_ids,
            selected_element_ids=selected_element_ids,
            require_all_visible=require_all_visible,
        )
        return view, False, False
    if selected == OPTION_DEDICATED_PLAN:
        view, created = get_or_create_dedicated_view(
            doc,
            "plan",
            preferred_level_id=preferred_level_id,
            logger=logger,
        )
        return view, created, False
    if selected == OPTION_DEDICATED_3D:
        view, created = get_or_create_dedicated_view(
            doc,
            "3d",
            preferred_level_id=preferred_level_id,
            logger=logger,
        )
        return view, created, True
    return None, False, False


def _apply_view_graphics(doc, view, element_ids, setup_dedicated, apply_overrides=True, logger=None):
    hidden_count = 0
    override_count = 0

    if (not setup_dedicated) and (not apply_overrides):
        return hidden_count, override_count

    tx = DB.Transaction(doc, "Circuit Element Finder: Prepare View")
    tx.Start()
    try:
        if setup_dedicated:
            hidden_count = hide_mep_categories(view, doc, logger=logger, include_mechanical_equipment=True)
        if apply_overrides:
            override_count = apply_selection_overrides(
                view,
                element_ids,
                line_color=DB.Color(240, 40, 40),
                line_weight=8,
                logger=logger,
            )
        tx.Commit()
    except Exception:
        try:
            tx.RollBack()
        except Exception:
            pass
        raise

    return hidden_count, override_count


def _apply_dedicated_3d_section_box(doc, view, element_ids, expanded_box=None, logger=None):
    if not isinstance(view, DB.View3D):
        return False
    expanded = expanded_box
    if expanded is None:
        bounds = get_combined_bounding_box(doc, element_ids, view=None)
        expanded = expand_bounding_box(bounds, padding_feet=4.0, minimum_half_extent=1.0)
    if expanded is None:
        return False

    tx = DB.Transaction(doc, "Circuit Element Finder: Set Section Box")
    tx.Start()
    try:
        try:
            view.IsSectionBoxActive = True
        except Exception:
            pass
        view.SetSectionBox(expanded)
        tx.Commit()
        return True
    except Exception as ex:
        try:
            tx.RollBack()
        except Exception:
            pass
        if logger:
            logger.debug("Section box update failed: {0}".format(ex))
        return False


def _collect_target_circuit_ids(doc):
    selection = list(revit.get_selection() or [])
    if selection:
        selected = []
        for element in list(selection or []):
            if isinstance(element, DB.Electrical.ElectricalSystem):
                selected.append(element)
    else:
        try:
            selected = list(eu.pick_circuits_from_list(doc, select_multiple=True) or [])
        except SystemExit:
            selected = []
    return [_id_value(circuit.Id) for circuit in list(selected or []) if isinstance(circuit, DB.Electrical.ElectricalSystem)]


def _collect_downstream_devices_strict(circuits):
    devices = []
    seen = set()
    zero_device_circuits = 0
    for circuit in list(circuits or []):
        circuit_id = _id_value(getattr(circuit, "Id", None))
        if circuit_id <= 0:
            continue
        try:
            connected = list(collect_circuit_targets(circuit, "device") or [])
        except Exception as ex:
            return None, "Failed to read connected devices for circuit {}: {}".format(circuit_id, ex)
        if not connected:
            zero_device_circuits += 1
        for element in list(connected or []):
            if element is None:
                continue
            element_id = getattr(element, "Id", None)
            element_value = _id_value(element_id)
            if element_value <= 0 or element_value in seen:
                continue
            seen.add(element_value)
            devices.append(element)
    return devices, "", int(zero_device_circuits)


def _show_and_select_elements(uidoc, elements, logger=None):
    element_ids = []
    seen = set()
    for element in list(elements or []):
        element_id = getattr(element, "Id", None)
        if not isinstance(element_id, DB.ElementId):
            continue
        element_value = _id_value(element_id)
        if element_value <= 0 or element_value in seen:
            continue
        seen.add(element_value)
        element_ids.append(element_id)
    if not element_ids:
        return False, False
    shown = show_elements(uidoc, element_ids, logger=logger)
    selected = set_revit_selection(elements, uidoc=uidoc)
    return bool(shown), bool(selected)


def run_circuited_device_finder(uidoc=None, logger=None):
    """Direct workflow: circuits -> downstream devices -> ShowElements + select."""
    log = logger or script.get_logger()

    active_uidoc = uidoc or getattr(revit, "uidoc", None)
    if active_uidoc is None:
        forms.alert("No active Revit UI document found.", title="Circuited Device Finder")
        return {"status": "error", "reason": "no_uidoc"}

    doc = getattr(active_uidoc, "Document", None) or getattr(revit, "doc", None)
    if doc is None:
        forms.alert("No active Revit document found.", title="Circuited Device Finder")
        return {"status": "error", "reason": "no_doc"}

    circuit_ids = [int(x) for x in list(_collect_target_circuit_ids(doc) or []) if int(x) > 0]
    if not circuit_ids:
        return {"status": "cancelled", "reason": "no_circuits"}
    circuits = []
    seen = set()
    for circuit_id in list(circuit_ids or []):
        if circuit_id in seen:
            continue
        seen.add(circuit_id)
        try:
            circuit = doc.GetElement(DB.ElementId(int(circuit_id)))
        except Exception:
            circuit = None
        if isinstance(circuit, DBE.ElectricalSystem):
            circuits.append(circuit)
    if not circuits:
        return {"status": "cancelled", "reason": "no_circuits"}

    devices, error_text, zero_device_circuits = _collect_downstream_devices_strict(circuits)
    if error_text:
        forms.alert(error_text, title="Circuited Device Finder")
        return {"status": "error", "reason": "device_collection_failed", "details": error_text}
    if not devices:
        forms.alert(
            "No downstream devices found.\n\n"
            "Circuits checked: {}\n"
            "Circuits with zero downstream devices: {}\n\n"
            "This can happen on spare/space/empty circuits or circuits with no connected load elements.".format(
                int(len(circuits or [])),
                int(zero_device_circuits or 0),
            ),
            title="Circuited Device Finder",
        )
        return {
            "status": "cancelled",
            "reason": "no_devices",
            "circuit_count": int(len(circuits or [])),
            "zero_device_circuits": int(zero_device_circuits or 0),
        }

    shown, selected = _show_and_select_elements(active_uidoc, devices, logger=log)
    if not shown:
        forms.alert("Could not show selected devices in model.", title="Circuited Device Finder")
        return {"status": "error", "reason": "show_elements_failed", "selected": bool(selected)}

    return {
        "status": "ok",
        "mode": "show_elements_only",
        "circuit_count": len(circuits),
        "zero_device_circuits": int(zero_device_circuits or 0),
        "device_count": len(devices),
        "zoomed": bool(shown),
        "selected": bool(selected),
    }


def run_circuit_element_finder(uidoc=None, logger=None):
    """Current active entrypoint for ribbon testing.

    TODO: Re-enable and optimize the 4 view-strategy methods:
    - OPTION_EXISTING_PLAN
    - OPTION_EXISTING_3D
    - OPTION_DEDICATED_PLAN
    - OPTION_DEDICATED_3D
    """
    return run_circuited_device_finder(uidoc=uidoc, logger=logger)
