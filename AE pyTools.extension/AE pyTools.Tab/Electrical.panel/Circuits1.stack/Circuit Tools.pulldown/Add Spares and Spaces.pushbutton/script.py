# -*- coding: utf-8 -*-
# Revit Python 2.7 – pyRevit / Revit API

from collections import defaultdict

import Autodesk.Revit.DB.Electrical as DBE
from Autodesk.Revit.DB import (
    Transaction, ElementId, SectionType
)
from pyrevit import revit, forms, script, DB
from pyrevit.compat import get_elementid_value_func

get_id_value = get_elementid_value_func()

doc  = revit.doc
uidoc = revit.uidoc
log  = script.get_logger()


# ---------------------------------------------------------------------------
# 0. Small helpers
# ---------------------------------------------------------------------------
def _get_param_str(el, name):
    """Case-insensitive lookup of a string-ish parameter; returns '' if missing."""
    if not el:
        return ''
    p = el.LookupParameter(name) or None
    if not p:
        # brute-force case-insensitive search
        try:
            for q in el.Parameters:
                try:
                    if q.Definition and q.Definition.Name and q.Definition.Name.lower() == name.lower():
                        p = q
                        break
                except Exception:
                    pass
        except Exception:
            pass
    if not p:
        return ''
    try:
        return (p.AsString() or p.AsValueString() or '') or ''
    except Exception:
        return ''

def _approx_equal(a, b, tol=1e-6):
    try:
        return abs(float(a) - float(b)) <= tol
    except Exception:
        return False

def _is_slot_locked(psv, row, col):
    """Best-effort read of lock state for a slot cell."""
    for attr in ('IsSlotLocked', 'GetLockSlot', 'IsCellLocked'):
        fn = getattr(psv, attr, None)
        if fn:
            try:
                return bool(fn(row, col))
            except Exception:
                pass
    return False

def _get_electrical_settings(doc):
    """Multiple API surfaces exist depending on Revit version."""
    try:
        return DBE.ElectricalSetting.GetElectricalSetting(doc)
    except Exception:
        try:
            return doc.Settings.ElectricalSettings
        except Exception:
            return None

def _get_default_circuit_rating(es):
    if not es:
        return None
    for attr in ('CircuitRating', ):
        try:
            return getattr(es, attr)
        except Exception:
            pass
    return None


# ---------------------------------------------------------------------------
# 1. Ask once how empty slots should be filled
# ---------------------------------------------------------------------------
def ask_fill_mode():
    mode = forms.CommandSwitchWindow.show(
        ['All Spare', 'All Space', 'Half Spare/Half Space'],
        title='Fill empty panel slots with...'
    )
    if not mode:
        forms.alert('Nothing chosen – cancelled.', exitscript=True)
    return mode  # str


def ask_remove_mode():
    mode = forms.CommandSwitchWindow.show(
        ['Spares only', 'Spaces only', 'Both'],
        title='Remove what?'
    )
    if not mode:
        forms.alert('Nothing chosen – cancelled.', exitscript=True)
    return mode


def ask_action():
    return forms.CommandSwitchWindow.show(
        ['Fill empty slots', 'Remove spares/spaces'],
        title='What do you want to do?'
    )


# ---------------------------------------------------------------------------
# 2. Collect panel-schedule views to process
# ---------------------------------------------------------------------------
class _ScheduleOption(object):
    """Wrapper so SelectFromList shows a name but returns the view object."""
    def __init__(self, view):
        self.view = view
    def __str__(self):
        return self.view.Name


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
    all_views = [v for v in DB.FilteredElementCollector(doc).OfClass(DBE.PanelScheduleView)
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
# 3. Scan a schedule and return { slot_num : [(row, col), …] } for *empty* cells
# ---------------------------------------------------------------------------
def gather_empty_cells(view):
    tbl  = view.GetTableData()
    body = tbl.GetSectionData(SectionType.Body)
    if not body:
        return {}

    max_slot = tbl.NumberOfSlots
    empties  = defaultdict(list)

    for row in range(body.NumberOfRows):
        active_slot = None
        cols_for_slot = []
        for col in range(body.NumberOfColumns):
            slot   = view.GetSlotNumberByCell(row, col)
            ckt_id = view.GetCircuitIdByCell(row, col)
            is_empty = (ckt_id == ElementId.InvalidElementId and 1 <= slot <= max_slot)

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
# 4. Reporting (shared by both fill & remove)
# ---------------------------------------------------------------------------
def report_results(title, rows):
    """
    rows = list of dicts, each including at least {'panel': name, ...other counters...}
    """
    out = script.get_output()
    out.set_title(title)
    out.print_md("# RESULTS\n")

    for idx, row in enumerate(rows, 1):
        name = row.get('panel', '(unknown)')
        out.print_md("## {}. {}".format(idx, name))
        for k in sorted(row.keys()):
            if k == 'panel':
                continue
            out.print_md("- {} : **{}**".format(k.replace('_', ' '), row[k]))
        if idx != len(rows):
            out.print_md("\n-----\n")
    out.show()


# ---------------------------------------------------------------------------
# 5. Fill schedules and return report rows (printing moved to report_results)
# ---------------------------------------------------------------------------
def fill_schedules(schedules, mode):
    rows = []  # [{'panel': name, 'open_slots_before': n, 'spares_added': x, 'spaces_added': y}]

    with Transaction(doc, 'Fill panel spares / spaces') as tx:
        tx.Start()

        for view in schedules:
            empty_map = gather_empty_cells(view)
            if not empty_map:
                rows.append({'panel': view.Name, 'open_slots_before': 0, 'spares_added': 0, 'spaces_added': 0})
                continue

            open_slots = len(empty_map)
            spare_cnt  = 0
            space_cnt  = 0

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
                                view.AddSpare(row, col)
                                spare_cnt += 1
                            else:
                                view.AddSpace(row, col)
                                space_cnt += 1
                            # ensure newly-added cell isn't locked
                            try:
                                view.SetLockSlot(row, col, 0)
                            except Exception:
                                pass
                            break
                        except Exception:
                            continue

            rows.append({'panel': view.Name,
                         'open_slots_before': open_slots,
                         'spares_added': spare_cnt,
                         'spaces_added': space_cnt})

        tx.Commit()

    return rows


# ---------------------------------------------------------------------------
# 6. Remove “removable” spares/spaces
# ---------------------------------------------------------------------------
def _is_removable_spare(psv, row, col, ckt, es, default_rating):
    """All Spare rules must be True."""
    # Rule: slot not locked
    if _is_slot_locked(psv, row, col):
        return False

    # Rules: Spare identity (either by circuit type OR by psv flag)
    is_spare_flag = False
    try:
        is_spare_flag = bool(psv.IsSpare(row, col))
    except Exception:
        pass

    ckt_type_is_spare = False
    try:
        # API enum name varies by version; try both
        ckt_type_is_spare = (getattr(ckt, 'CircuitType', None) in (
            getattr(DBE, 'ElectricalCircuitType', None) and DBE.ElectricalCircuitType.Spare or None,
            getattr(DBE, 'CircuitType', None) and DBE.CircuitType.Spare or None
        ))
    except Exception:
        pass

    if not (is_spare_flag or ckt_type_is_spare):
        return False

    # Rule: LoadName == "spare" (case-insensitive)
    try:
        loadname = (ckt.LoadName or '').strip().lower()
    except Exception:
        loadname = ''
    if loadname != 'spare':
        return False

    # Rule: schedule circuit notes empty
    notes = _get_param_str(ckt, 'Schedule Circuit Notes').strip()
    if notes:
        return False

    # Rule: ApparentLoad == 0
    try:
        if abs(float(ckt.ApparentLoad)) > 1e-6:
            return False
    except Exception:
        return False

    # Rule: Rating == Electrical Settings default rating
    try:
        ckt_rating = getattr(ckt, 'Rating', None)
        if default_rating is None or ckt_rating is None or not _approx_equal(ckt_rating, default_rating):
            return False
    except Exception:
        return False

    return True


def _is_removable_space(psv, row, col, ckt):
    """All Space rules must be True."""
    # Rule: slot not locked
    if _is_slot_locked(psv, row, col):
        return False

    # Rule: identity (space)
    is_space_flag = False
    try:
        is_space_flag = bool(psv.IsSpace(row, col))
    except Exception:
        pass

    ckt_type_is_space = False
    try:
        ckt_type_is_space = (getattr(ckt, 'CircuitType', None) in (
            getattr(DBE, 'ElectricalCircuitType', None) and DBE.ElectricalCircuitType.Space or None,
            getattr(DBE, 'CircuitType', None) and DBE.CircuitType.Space or None
        ))
    except Exception:
        pass

    try:
        loadname = (ckt.LoadName or '').strip().lower()
    except Exception:
        loadname = ''

    return (is_space_flag or ckt_type_is_space) and (loadname == 'space')


def remove_spares_spaces(schedules, remove_mode):
    """
    Remove spares and/or spaces that satisfy 'Removable Rules'.
    remove_mode: 'Spares only' | 'Spaces only' | 'Both'
    """
    rows = []  # [{'panel': name, 'spares_removed': x, 'spaces_removed': y}]

    es = _get_electrical_settings(doc)
    default_rating = _get_default_circuit_rating(es)

    with Transaction(doc, 'Remove panel spares / spaces') as tx:
        tx.Start()

        for psv in schedules:
            tbl  = psv.GetTableData()
            body = tbl.GetSectionData(SectionType.Body)
            if not body:
                rows.append({'panel': psv.Name, 'spares_removed': 0, 'spaces_removed': 0})
                continue

            processed_slots = set()
            sp_removed = 0
            sc_removed = 0

            for row in range(body.NumberOfRows):
                for col in range(body.NumberOfColumns):
                    try:
                        slot = psv.GetSlotNumberByCell(row, col)
                    except Exception:
                        slot = 0
                    if slot < 1 or slot > tbl.NumberOfSlots:
                        continue
                    if slot in processed_slots:
                        continue

                    # detect spare/space at this cell
                    is_spare = False
                    is_space = False
                    try:
                        is_spare = bool(psv.IsSpare(row, col))
                    except Exception:
                        pass
                    try:
                        is_space = bool(psv.IsSpace(row, col))
                    except Exception:
                        pass
                    if not (is_spare or is_space):
                        continue  # not a spare/space cell for this slot

                    # now we will process this slot exactly once
                    processed_slots.add(slot)

                    # fetch circuit if any (spares/spaces should have an ElectricalSystem behind them)
                    ckt_id = psv.GetCircuitIdByCell(row, col)
                    ckt = doc.GetElement(ckt_id) if (ckt_id and ckt_id != ElementId.InvalidElementId) else None

                    # Removal decision per mode and rules
                    if (remove_mode in ('Spares only', 'Both')) and is_spare:
                        if ckt and _is_removable_spare(psv, row, col, ckt, es, default_rating):
                            try:
                                psv.RemoveSpare(row, col)
                                sp_removed += 1
                            except Exception as e:
                                log.debug('RemoveSpare failed on "{}" (r{},c{}): {}'.format(psv.Name, row, col, e))
                        continue  # don't attempt 'space' on same slot

                    if (remove_mode in ('Spaces only', 'Both')) and is_space:
                        if ckt and _is_removable_space(psv, row, col, ckt):
                            try:
                                psv.RemoveSpace(row, col)
                                sc_removed += 1
                            except Exception as e:
                                log.debug('RemoveSpace failed on "{}" (r{},c{}): {}'.format(psv.Name, row, col, e))
                        continue

            rows.append({'panel': psv.Name,
                         'spares_removed': sp_removed,
                         'spaces_removed': sc_removed})

        tx.Commit()

    return rows


# ---------------------------------------------------------------------------
# 7. Main
# ---------------------------------------------------------------------------
def main():
    scheds = collect_schedules_to_process()
    action = ask_action()

    if action == 'Fill empty slots':
        fill_mode = ask_fill_mode()
        rows = fill_schedules(scheds, fill_mode)
        report_results('Panel-Schedule Fill Results', rows)
    else:
        rm_mode = ask_remove_mode()
        rows = remove_spares_spaces(scheds, rm_mode)
        report_results('Panel-Schedule Removal Results', rows)


if __name__ == '__main__':
    main()
