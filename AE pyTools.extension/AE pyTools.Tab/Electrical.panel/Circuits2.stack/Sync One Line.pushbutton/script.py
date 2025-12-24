# -*- coding: utf-8 -*-
import os
from pyrevit import revit, DB, script, forms
import System
from Autodesk.Revit.Exceptions import OperationCanceledException
from CEDElectrical.Domain.one_line_sync import OneLineSyncService
from pyrevitmep.event import CustomizableEvent
import UIClasses.SyncOneLineWindow as sync_ui
from UIClasses.SyncOneLineWindow import SyncOneLineWindow, SyncOneLineListItem, status_symbol

logger = script.get_logger()

XAML_PATH = os.path.join(os.path.dirname(sync_ui.__file__), "SyncOneLineWindow.xaml")
_ONE_LINE_WINDOW = None
_ONE_LINE_EVENT = None
ENABLE_ONE_LINE_LOG = True


def get_element_label(assoc):
    elem = assoc.model_elem
    if assoc.kind == "circuit":
        try:
            panel = elem.BaseEquipment
            panel_name = panel.Name if panel else ""
        except Exception:
            panel_name = ""
        try:
            cnum = elem.CircuitNumber
        except Exception:
            cnum = ""
        return "{} {}".format(panel_name, cnum).strip()
    if assoc.kind == "device":
        try:
            symbol = elem.Symbol
            if symbol and symbol.Family:
                return "{}: {} - {}".format(symbol.Family.Name, symbol.Name, elem.Name)
        except Exception:
            pass
    try:
        return elem.Name
    except Exception:
        return "Element {}".format(elem.Id.IntegerValue)


def sort_key_flat(assoc):
    order = {"circuit": 0, "panel": 1, "device": 2}
    category_order = order.get(assoc.kind, 3)
    if assoc.kind == "circuit":
        panel = ""
        cnum = ""
        try:
            base = assoc.model_elem.BaseEquipment
            panel = base.Name if base else ""
        except Exception:
            panel = ""
        try:
            cnum = assoc.model_elem.CircuitNumber
        except Exception:
            cnum = ""
        return (category_order, panel, get_circuit_sort_key(assoc.model_elem), str(cnum))
    if assoc.kind == "panel":
        return (category_order, get_element_label(assoc))
    return (category_order, get_element_label(assoc))


def get_kind_symbol(assoc):
    if assoc.kind == "panel":
        return u"\u25A0"
    if assoc.kind == "device":
        return u"\u25CF"
    if assoc.kind == "circuit":
        if getattr(assoc, "is_spare", False):
            return u"\u25B7"
        return u"\u25B6"
    return u"\u25A1"


def build_list_items(associations, tree_order):
    items = []
    assoc_to_item = {}
    for assoc in associations:
        label = get_element_label(assoc)
        display = "{} (Id {})".format(label, assoc.model_elem.Id.IntegerValue)
        symbol, brush = status_symbol(assoc.status)
        kind_symbol = get_kind_symbol(assoc)
        item = SyncOneLineListItem(assoc, display, display, symbol, brush, kind_symbol)
        items.append(item)
        assoc_to_item[assoc] = item

    for assoc, indent in tree_order:
        if assoc in assoc_to_item:
            base_text = assoc_to_item[assoc].base_text
            assoc_to_item[assoc].tree_text = "{}{}".format("    " * indent, base_text)

    return items, assoc_to_item


def build_ordered_items(tree_order, assoc_to_item):
    ordered = []
    for assoc, _indent in tree_order:
        item = assoc_to_item.get(assoc)
        if item:
            ordered.append(item)
    return ordered


def set_selection(ids):
    id_list = System.Collections.Generic.List[DB.ElementId]()
    for elem_id in ids:
        id_list.Add(elem_id)
    revit.uidoc.Selection.SetElementIds(id_list)


def build_placement_points(start_point, end_point, count):
    if count <= 1:
        return [start_point]

    dx = end_point.X - start_point.X
    dy = end_point.Y - start_point.Y
    dz = end_point.Z - start_point.Z
    tol = 1e-6
    if abs(dx) >= abs(dy):
        step_x = dx / float(count - 1)
        step_y = 0.0
    else:
        step_x = 0.0
        step_y = dy / float(count - 1)
    step_z = dz / float(count - 1) if abs(dz) > tol else 0.0

    points = []
    for idx in range(count):
        points.append(DB.XYZ(start_point.X + step_x * idx,
                             start_point.Y + step_y * idx,
                             start_point.Z + step_z * idx))
    return points


def get_circuit_sort_key(branch):
    val = None
    try:
        val = branch.CircuitNumber
    except Exception:
        try:
            val = branch.circuit_number
        except Exception:
            val = ""
    try:
        return int(val)
    except Exception:
        return str(val)


def is_spare_or_space(system):
    try:
        cnum = system.CircuitNumber or ""
        return "spare" in cnum.lower() or "space" in cnum.lower()
    except Exception:
        return False


def get_assigned_systems(equipment):
    try:
        mep = equipment.MEPModel
        if not mep:
            return []
        return list(mep.GetAssignedElectricalSystems())
    except Exception:
        return []


def get_connected_systems(equipment):
    try:
        mep = equipment.MEPModel
        if not mep:
            return []
        return list(mep.GetElectricalSystems())
    except Exception:
        return []


def is_root_equipment(equipment):
    assigned = get_assigned_systems(equipment)
    if not assigned:
        return False
    connected = get_connected_systems(equipment)
    return len(assigned) == len(connected)


def build_tree_order(associations, doc):
    assoc_by_id = {}
    for assoc in associations:
        assoc_by_id[assoc.model_elem.Id.IntegerValue] = assoc

    ordered = []
    visited = set()

    def add_assoc(assoc, indent):
        if not assoc:
            return
        assoc_id = assoc.model_elem.Id.IntegerValue
        if assoc_id in visited:
            return
        ordered.append((assoc, indent))
        visited.add(assoc_id)

    def walk_circuit(system, indent):
        assoc = assoc_by_id.get(system.Id.IntegerValue)
        if assoc:
            assoc.is_spare = is_spare_or_space(system)
        add_assoc(assoc, indent)
        elements = []
        try:
            elements = list(system.Elements)
        except Exception:
            elements = []
        for elem in elements:
            if elem.Category and int(elem.Category.Id.IntegerValue) == int(DB.BuiltInCategory.OST_ElectricalEquipment):
                walk_equipment(elem, indent + 1)
            else:
                add_assoc(assoc_by_id.get(elem.Id.IntegerValue), indent + 1)

    def walk_equipment(equipment, indent):
        assoc = assoc_by_id.get(equipment.Id.IntegerValue)
        add_assoc(assoc, indent)
        systems = get_assigned_systems(equipment)
        for system in sorted(systems, key=get_circuit_sort_key):
            walk_circuit(system, indent + 1)

    equipment_assocs = [assoc for assoc in associations if assoc.kind == "panel"]
    root_equipment = []
    for assoc in equipment_assocs:
        try:
            if is_root_equipment(assoc.model_elem):
                root_equipment.append(assoc.model_elem)
        except Exception:
            continue

    for equip in sorted(root_equipment, key=lambda e: e.Name or ""):
        walk_equipment(equip, 0)

    circuit_assocs = [assoc for assoc in associations if assoc.kind == "circuit"]
    for assoc in sorted(circuit_assocs, key=lambda a: get_circuit_sort_key(a.model_elem)):
        add_assoc(assoc, 0)

    for assoc in associations:
        if assoc.model_elem.Id.IntegerValue not in visited:
            ordered.append((assoc, 0))

    return ordered


def ensure_detail_view(view):
    if view.ViewType in [DB.ViewType.DraftingView, DB.ViewType.Detail, DB.ViewType.FloorPlan,
                         DB.ViewType.CeilingPlan, DB.ViewType.Section, DB.ViewType.Elevation]:
        return True
    return False


def refresh_statuses(service, items):
    for item in items:
        item.association.status = service.compute_status(item.association)
        symbol, brush = status_symbol(item.association.status)
        item.status_symbol = symbol
        item.status_brush = brush


def log_action(message):
    if ENABLE_ONE_LINE_LOG:
        output = script.get_output()
        output.print_md("* {}".format(message))


def execute_sync(window, service, doc, items):
    log_action("Sync Selected clicked.")
    selected_associations = window.get_selected_associations()
    if not selected_associations:
        forms.alert("Please select at least one element.")
        return

    t = DB.Transaction(doc, "Sync One-Line Detail Items")
    t.Start()
    updated = service.sync_associations(selected_associations)
    t.Commit()

    refresh_statuses(service, items)
    window.refresh_items()

    warnings = service.get_link_warnings(selected_associations)
    if warnings:
        output = script.get_output()
        output.print_md("## Sync One-Line Warnings")
        for warning in warnings:
            output.print_md("* {}".format(warning))
    else:
        forms.alert("Synced {} element(s).".format(updated))


def execute_create(window, service, doc, view, items):
    log_action("Create Detail Items clicked.")
    if not ensure_detail_view(view):
        forms.alert("Active view must support detail items.")
        return

    selected_associations = window.get_selected_associations()
    if not selected_associations:
        forms.alert("Please select at least one element.")
        return

    detail_symbol = window.get_selected_detail_symbol()
    if not detail_symbol:
        forms.alert("Select a detail item family and type before creating.")
        return

    tag_symbol = window.get_selected_tag_symbol()

    try:
        window.Hide()
        start_point = revit.uidoc.Selection.PickPoint("Pick start point for detail items")
        end_point = revit.uidoc.Selection.PickPoint("Pick end point for detail items")
    except OperationCanceledException:
        window.Show()
        window.Activate()
        return
    finally:
        window.Show()
        window.Activate()

    associations_to_create = [assoc for assoc in selected_associations if assoc.detail_elem is None]
    if not associations_to_create:
        forms.alert("Selected elements already have detail items.")
        return

    points = build_placement_points(start_point, end_point, len(associations_to_create))

    t = DB.Transaction(doc, "Create One-Line Detail Items")
    t.Start()
    created = service.create_detail_items(associations_to_create, detail_symbol, view, points, tag_symbol)
    updated = service.sync_associations(created)
    t.Commit()

    refresh_statuses(service, items)
    window.refresh_items()

    warnings = service.get_link_warnings(created)
    if warnings:
        output = script.get_output()
        output.print_md("## Create Detail Items Warnings")
        for warning in warnings:
            output.print_md("* {}".format(warning))
    else:
        forms.alert("Created {} detail item(s) and synced {} element(s).".format(len(created), updated))


def execute_select(window, model=True):
    action = "Select Model" if model else "Select Detail"
    log_action("{} clicked.".format(action))
    items = window.ElementsList.SelectedItems
    if not items or items.Count == 0:
        return
    assoc = items[0].association
    if model:
        set_selection([assoc.model_elem.Id])
    else:
        if assoc.detail_elem:
            set_selection([assoc.detail_elem.Id])


def execute_update_details(window, service, item):
    log_action("Selection changed.")
    if not item:
        window.set_detail_panel("(none)", "-", "-", [])
        return
    assoc = item.association
    detail_id = "-" if not assoc.detail_elem else str(assoc.detail_elem.Id.IntegerValue)
    model_id = str(assoc.model_elem.Id.IntegerValue)
    label = get_element_label(assoc)
    summary = []
    comparisons = service.compare_values(assoc)
    for comp in comparisons:
        status = "OK" if comp["match"] else "Outdated"
        summary.append("{}: {} | Model='{}' Detail='{}'".format(
            status, comp["param"], comp["model"], comp["detail"]))
    window.set_detail_panel(label, model_id, detail_id, summary)


def main():
    doc = revit.doc
    view = revit.active_view

    service = OneLineSyncService(doc)
    associations = service.build_associations()

    if not associations:
        forms.alert("No circuits, panels, or devices found.")
        return

    tree_order = build_tree_order(associations, doc)
    sorted_associations = sorted(associations, key=sort_key_flat)
    list_items, assoc_to_item = build_list_items(associations, tree_order)
    flat_items = [assoc_to_item[assoc] for assoc in sorted_associations if assoc in assoc_to_item]
    tree_items = build_ordered_items(tree_order, assoc_to_item)

    detail_symbols = service.collect_detail_symbols()
    tag_symbols = service.collect_tag_symbols()

    global _ONE_LINE_WINDOW
    global _ONE_LINE_HANDLER
    global _ONE_LINE_EVENT

    customizable_event = CustomizableEvent()
    customizable_event.logger = logger
    _ONE_LINE_EVENT = customizable_event

    window = SyncOneLineWindow(XAML_PATH, flat_items, tree_items, detail_symbols, tag_symbols,
                               on_sync=lambda: customizable_event.raise_event(execute_sync, window, service, doc, list_items),
                               on_create=lambda: customizable_event.raise_event(execute_create, window, service, doc, view, list_items),
                               on_select_model=lambda: customizable_event.raise_event(execute_select, window, True),
                               on_select_detail=lambda: customizable_event.raise_event(execute_select, window, False),
                               on_selection_changed=lambda item: customizable_event.raise_event(execute_update_details, window, service, item))
    _ONE_LINE_WINDOW = window
    window.Show()


if __name__ == "__main__":
    main()
