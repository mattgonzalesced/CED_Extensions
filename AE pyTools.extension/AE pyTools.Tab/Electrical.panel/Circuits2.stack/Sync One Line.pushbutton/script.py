# -*- coding: utf-8 -*-
import os
from pyrevit import revit, DB, script, forms
from Autodesk.Revit.Exceptions import OperationCanceledException
from CEDElectrical.Domain.one_line_sync import OneLineSyncService
from CEDElectrical.Domain.one_line_tree import build_system_tree
import UIClasses.SyncOneLineWindow as sync_ui
from UIClasses.SyncOneLineWindow import SyncOneLineWindow, SyncOneLineListItem, status_symbol

logger = script.get_logger()

XAML_PATH = os.path.join(os.path.dirname(sync_ui.__file__), "SyncOneLineWindow.xaml")


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
    try:
        return elem.Name
    except Exception:
        return "Element {}".format(elem.Id.IntegerValue)


def build_list_items(associations, tree_order):
    items = []
    assoc_to_item = {}
    for assoc in associations:
        label = get_element_label(assoc)
        kind_label = assoc.kind.capitalize()
        display = "[{}] {} (Id {})".format(kind_label, label, assoc.model_elem.Id.IntegerValue)
        symbol, brush = status_symbol(assoc.status)
        item = SyncOneLineListItem(assoc, display, display, symbol, brush)
        items.append(item)
        assoc_to_item[assoc] = item

    for assoc, indent in tree_order:
        if assoc in assoc_to_item:
            base_text = assoc_to_item[assoc].base_text
            assoc_to_item[assoc].tree_text = "{}{}".format("    " * indent, base_text)

    return items


def get_circuit_sort_key(branch):
    val = branch.circuit_number or ""
    try:
        return int(val)
    except Exception:
        return str(val)


def build_tree_order(associations, doc):
    assoc_by_id = {}
    for assoc in associations:
        assoc_by_id[assoc.model_elem.Id.IntegerValue] = assoc

    ordered = []
    visited = set()
    tree = build_system_tree(doc)

    def add_assoc(assoc, indent):
        if not assoc:
            return
        assoc_id = assoc.model_elem.Id.IntegerValue
        if assoc_id in visited:
            return
        ordered.append((assoc, indent))
        visited.add(assoc_id)

    def walk_node(node, indent):
        add_assoc(assoc_by_id.get(node.element_id.IntegerValue), indent)
        for branch in sorted(node.downstream, key=get_circuit_sort_key):
            add_assoc(assoc_by_id.get(branch.element_id.IntegerValue), indent + 1)
            system = branch.system
            if hasattr(system, "Elements"):
                for elem in list(system.Elements):
                    if elem.Category and int(elem.Category.Id.IntegerValue) == int(DB.BuiltInCategory.OST_ElectricalEquipment):
                        child_node = tree.get_node(elem.Id)
                        if child_node:
                            walk_node(child_node, indent + 2)
                    else:
                        add_assoc(assoc_by_id.get(elem.Id.IntegerValue), indent + 2)

    for root in sorted(tree.root_nodes, key=lambda n: n.panel_name or ""):
        walk_node(root, 0)

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


def main():
    doc = revit.doc
    view = revit.active_view

    service = OneLineSyncService(doc)
    associations = service.build_associations()

    if not associations:
        forms.alert("No circuits, panels, or devices found.")
        return

    tree_order = build_tree_order(associations, doc)
    list_items = build_list_items(associations, tree_order)

    detail_symbols = service.collect_detail_symbols()
    tag_symbols = service.collect_tag_symbols()

    def on_sync():
        selected_associations = window.get_selected_associations()
        if not selected_associations:
            forms.alert("Please select at least one element.")
            return

        t = DB.Transaction(doc, "Sync One-Line Detail Items")
        t.Start()
        updated = service.sync_associations(selected_associations)
        t.Commit()

        refresh_statuses(service, list_items)
        window.refresh_items()

        warnings = service.get_link_warnings(selected_associations)
        if warnings:
            forms.alert("Sync completed with warnings:\n\n- " + "\n- ".join(warnings))
        else:
            forms.alert("Synced {} element(s).".format(updated))

    def on_create():
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
            base_point = revit.uidoc.Selection.PickPoint("Pick insertion point for detail items")
        except OperationCanceledException:
            window.Show()
            return
        finally:
            window.Show()

        associations_to_create = [assoc for assoc in selected_associations if assoc.detail_elem is None]
        if not associations_to_create:
            forms.alert("Selected elements already have detail items.")
            return

        t = DB.Transaction(doc, "Create One-Line Detail Items")
        t.Start()
        created = service.create_detail_items(associations_to_create, detail_symbol, view, base_point, tag_symbol)
        updated = service.sync_associations(created)
        t.Commit()

        refresh_statuses(service, list_items)
        window.refresh_items()

        warnings = service.get_link_warnings(created)
        if warnings:
            forms.alert("Created {} detail item(s) with warnings:\n\n- {}".format(
                len(created), "\n- ".join(warnings)))
        else:
            forms.alert("Created {} detail item(s) and synced {} element(s).".format(len(created), updated))

    window = SyncOneLineWindow(XAML_PATH, list_items, detail_symbols, tag_symbols,
                               on_sync=on_sync, on_create=on_create)
    window.ShowDialog()


if __name__ == "__main__":
    main()
