# -*- coding: utf-8 -*-
# Revit Python 2.7  – pyRevit / Revit API

from collections import defaultdict

import Autodesk.Revit.DB.Electrical as DBE
from Autodesk.Revit.DB import (
    FilteredElementCollector, Transaction, ElementId,
    SectionType
)
from pyrevit import revit, forms, script
from pyrevit.compat import get_elementid_value_func

get_id_value = get_elementid_value_func()

doc = revit.doc
uidoc = revit.uidoc
out = script.get_output()
log = script.get_logger()


# ---------------------------------------------------------------------------
# 1. Ask once how empty slots should be filled
# ---------------------------------------------------------------------------
def ask_fill_mode():
    mode = forms.CommandSwitchWindow.show(
        ['All Spare', 'All Space', 'Half Spare/Half Space'],
        title='Fill empty panel slots with...')
    if not mode:
        forms.alert('Nothing chosen – cancelled.', exitscript=True)
    return mode  # str


# ---------------------------------------------------------------------------
# 2. Collect panel‑schedule views to process
# ---------------------------------------------------------------------------
class _ScheduleOption(object):
    """Wrapper so SelectFromList shows a name but returns the view object."""

    def __init__(self, view): self.view = view

    def __str__(self):       return self.view.Name


def _schedules_from_selection(elements):
    found, skipped = [], defaultdict(int)
    for el in elements:
        if isinstance(el, DBE.PanelScheduleSheetInstance):
            v = doc.GetElement(el.ScheduleId)
            if isinstance(v, DBE.PanelScheduleView):
                found.append(v)
        else:
            cat = el.Category.Name if el.Category else 'Unknown'
            skipped[cat] += 1

    for cat, cnt in skipped.items():
        log.warning('{} “{}” element(s) skipped'.format(cnt, cat))

    # remove duplicates
    uniq = {get_id_value(v.Id): v for v in found}.values()
    return list(uniq)


def _prompt_for_schedules():
    all_views = [v for v in FilteredElementCollector(doc)
    .OfClass(DBE.PanelScheduleView)
                 if not v.IsTemplate]
    if not all_views:
        forms.alert('No panel schedules in this model.', exitscript=True)

    picked = forms.SelectFromList.show(
        [_ScheduleOption(v) for v in sorted(all_views, key=lambda x: x.Name)],
        title='Choose panel schedules', multiselect=True)
    if not picked:
        forms.alert('Nothing selected – cancelled.', exitscript=True)
    return [p.view for p in picked]


def collect_schedules_to_process():
    # a) active view
    av = uidoc.ActiveView
    if isinstance(av, DBE.PanelScheduleView):
        return [av]

    # b) graphics selected on sheet
    sel = revit.get_selection()
    if sel:
        views = _schedules_from_selection(sel.elements)
        if views:
            return views

    # c) let user pick
    return _prompt_for_schedules()


# ---------------------------------------------------------------------------
# 3. Scan a schedule and return { slot_num : [(row, col), …] }
# ---------------------------------------------------------------------------
def gather_empty_cells(view):
    tbl = view.GetTableData()
    body = tbl.GetSectionData(SectionType.Body)
    if not body:
        return {}

    max_slot = tbl.NumberOfSlots
    empties = defaultdict(list)

    for row in range(body.NumberOfRows):
        active_slot = None
        cols_for_slot = []
        for col in range(body.NumberOfColumns):
            slot = view.GetSlotNumberByCell(row, col)
            ckt_id = view.GetCircuitIdByCell(row, col)
            is_empty = (ckt_id == ElementId.InvalidElementId
                        and 1 <= slot <= max_slot)

            if is_empty and slot == active_slot:
                cols_for_slot.append(col)
            else:
                if active_slot and cols_for_slot:
                    empties[active_slot].extend((row, c) for c in cols_for_slot)
                active_slot = slot if is_empty else None
                cols_for_slot = [col] if is_empty else []

        if active_slot and cols_for_slot:
            empties[active_slot].extend((row, c) for c in cols_for_slot)

    return empties


# ---------------------------------------------------------------------------
# 4. Fill schedules according to the chosen mode and print summary
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# 4. Fill schedules and build report  (console opens *after* commit)
# ---------------------------------------------------------------------------
def fill_schedules(schedules, mode):
    results = []  # [(panelName, open, spare, space)]

    with Transaction(doc, 'Fill panel spares / spaces') as tx:
        tx.Start()

        for view in schedules:
            empty_map = gather_empty_cells(view)
            if not empty_map:
                results.append((view.Name, 0, 0, 0))
                continue

            open_slots = len(empty_map)
            spare_cnt = 0
            space_cnt = 0

            slot_items = sorted(empty_map.items())

            if mode == 'All Spare':
                work = [(True, slot_items)]
            elif mode == 'All Space':
                work = [(False, slot_items)]
            else:
                half = len(slot_items) // 2
                work = [(True, slot_items[:half]),
                        (False, slot_items[half:])]

            for want_spare, chunk in work:
                for slot, cells in chunk:
                    for row, col in cells:
                        try:
                            if want_spare:
                                view.AddSpare(row, col);
                                spare_cnt += 1
                            else:
                                view.AddSpace(row, col);
                                space_cnt += 1
                            view.SetLockSlot(row, col, 0)
                            break
                        except Exception:
                            continue

            results.append((view.Name, open_slots, spare_cnt, space_cnt))

        tx.Commit()  # ------------------

    # ----------- console output happens only *after* commit -----------------
    out = script.get_output()
    out.set_title("Panel‑Schedule Fill Results")
    out.print_md("# RESULTS\n")

    for idx, (name, open_slots, spare_cnt, space_cnt) in enumerate(results, 1):
        out.print_md("## {}. {}".format(idx, name))
        out.print_md("- open slots before : **{}**".format(open_slots))
        out.print_md("- spares added      : **{}**".format(spare_cnt))
        out.print_md("- spaces added      : **{}**".format(space_cnt))
        if idx != len(results):
            out.print_md("\n-----\n")

    out.show()  # pops console


# ---------------------------------------------------------------------------
def main():
    scheds = collect_schedules_to_process()
    mode = ask_fill_mode()
    fill_schedules(scheds, mode)


if __name__ == '__main__':
    main()
