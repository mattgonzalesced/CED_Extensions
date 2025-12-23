# -*- coding: utf-8 -*-
import os
from pyrevit import revit, DB, script, forms
from CEDElectrical.Domain.one_line_sync import OneLineSyncService
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


def build_list_items(associations):
    items = []
    for assoc in associations:
        label = get_element_label(assoc)
        kind_label = assoc.kind.capitalize()
        display = "[{}] {} (Id {})".format(kind_label, label, assoc.model_elem.Id.IntegerValue)
        symbol, brush = status_symbol(assoc.status)
        items.append(SyncOneLineListItem(assoc, display, symbol, brush))
    return items


def ensure_detail_view(view):
    if view.ViewType in [DB.ViewType.DraftingView, DB.ViewType.Detail, DB.ViewType.FloorPlan,
                         DB.ViewType.CeilingPlan, DB.ViewType.Section, DB.ViewType.Elevation]:
        return True
    return False


def main():
    doc = revit.doc
    view = revit.active_view

    service = OneLineSyncService(doc)
    associations = service.build_associations()

    if not associations:
        forms.alert("No circuits, panels, or devices found.")
        return

    list_items = build_list_items(associations)

    detail_symbols = service.collect_detail_symbols()
    tag_symbols = service.collect_tag_symbols()

    window = SyncOneLineWindow(XAML_PATH, list_items, detail_symbols, tag_symbols)
    window.ShowDialog()

    action = window.requested_action
    if not action:
        return

    selected_associations = window.get_selected_associations()
    if not selected_associations:
        forms.alert("Please select at least one element.")
        return

    if action == "create":
        if not ensure_detail_view(view):
            forms.alert("Active view must support detail items.")
            return

        detail_symbol = window.get_selected_detail_symbol()
        if not detail_symbol:
            forms.alert("Select a detail item family and type before creating.")
            return

        tag_symbol = window.get_selected_tag_symbol()
        panel_associations = [assoc for assoc in selected_associations if assoc.kind == "panel"]
        if not panel_associations:
            forms.alert("Select panel equipment to create detail items.")
            return

        t = DB.Transaction(doc, "Create One-Line Detail Items")
        t.Start()
        created = service.create_detail_items(panel_associations, detail_symbol, view, tag_symbol)
        updated = service.sync_associations(created)
        t.Commit()

        forms.alert("Created {} detail item(s) and synced {} panel(s).".format(len(created), updated))
        return

    if action == "sync":
        t = DB.Transaction(doc, "Sync One-Line Detail Items")
        t.Start()
        updated = service.sync_associations(selected_associations)
        t.Commit()

        forms.alert("Synced {} element(s).".format(updated))


if __name__ == "__main__":
    main()
