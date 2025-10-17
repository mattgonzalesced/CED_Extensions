# -*- coding: utf-8 -*-
# Revit 2024 / IronPython 2.7 (pyRevit)
# Swap OLD base views/schedules with their NEW "<name>1" duplicates while preserving sheet placements.
# - Views & Legends handled via Viewport
# - Schedules handled via per-sheet scan using ViewSheet.GetAllPlacedScheduleInstances / GetSchedulePlacedInstances
#   with static/fallbacks for other builds
# - Safety: if a schedule has no detectable placements, archive OLD (don't delete) so sheets don't vanish.

from __future__ import print_function
import re, traceback

from Autodesk.Revit.DB import (
    FilteredElementCollector, View, Viewport, ViewSheet, BuiltInCategory,
    BuiltInParameter, Transaction, TransactionGroup, XYZ, ElementId,
    ViewSchedule, ScheduleSheetInstance
)
from pyrevit import script

logger = script.get_logger()
output = script.get_output()

# -----------------------------
# CONFIG
# -----------------------------
DRY_RUN = False
ONLY_FIRST_PAIR = False
ENABLE_SCHEDULES = True
EXTRA_DIAG = True
SAFE_SKIP_DELETE_IF_NO_PLACEMENTS = True

_SUFFIX_PATTERNS = [
    re.compile(r'^(?P<base>.+?)(?P<num>\d+)$'),         # "Name1"
]

# -----------------------------
# Helpers (names/types)
# -----------------------------
def strip_numeric_suffix(name):
    for pat in _SUFFIX_PATTERNS:
        m = pat.match(name)
        if m:
            try:
                return m.group('base'), int(m.group('num'))
            except:
                return m.group('base'), None
    return name, None

def is_schedule(v):
    return isinstance(v, ViewSchedule)

def is_placeable_view(v):
    return (v is not None) and (not v.IsTemplate)

# -----------------------------
# Title on Sheet
# -----------------------------
def get_title_on_sheet(view):
    p = view.LookupParameter("Title on Sheet")
    if p:
        try:
            return p.AsString()
        except:
            return None
    return None

def set_title_on_sheet(view, value):
    if not value:
        return
    p = view.LookupParameter("Title on Sheet")
    if p and not p.IsReadOnly:
        try:
            p.Set(value)
        except Exception as e:
            logger.warning("Could not set 'Title on Sheet' on '{}': {}".format(view.Name, e))

# -----------------------------
# Viewports (views & legends)
# -----------------------------
def collect_viewport_placements(doc, view_id):
    out = []
    vps = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Viewports).WhereElementIsNotElementType()
    for vp in vps:
        if vp.ViewId == view_id:
            sheet = doc.GetElement(vp.SheetId)
            try:
                center = vp.GetBoxCenter()
            except:
                outln = vp.GetBoxOutline()
                mi, ma = outln.MinimumPoint, outln.MaximumPoint
                center = XYZ((mi.X+ma.X)/2.0, (mi.Y+ma.Y)/2.0, 0.0)
            out.append(dict(
                sheet_id=vp.SheetId,
                sheet_no=sheet.SheetNumber,
                sheet_name=sheet.Name,
                center=center,
                vp_type_id=vp.GetTypeId(),
                detail_number=(vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER).AsString()
                               if vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER) else None)
            ))
    return out

def can_add_view_to_sheet(doc, sheet, view):
    try:
        return Viewport.CanAddViewToSheet(doc, sheet.Id, view.Id)
    except:
        return False

def place_view_on_sheet(doc, sheet, view, center_xyz, vp_type_id):
    if not can_add_view_to_sheet(doc, sheet, view):
        logger.warning("    - View '{}' cannot be added to sheet {}.".format(view.Name, sheet.SheetNumber))
        return None
    vp = Viewport.Create(doc, sheet.Id, view.Id, center_xyz)
    try:
        if vp_type_id and vp.GetTypeId() != vp_type_id:
            vp.ChangeTypeId(vp_type_id)
    except Exception as e:
        logger.warning("    - Could not set viewport type on {}: {}".format(sheet.SheetNumber, e))
    return vp

def existing_detail_numbers_on_sheet(doc, sheet):
    nums = set()
    vps = FilteredElementCollector(doc, sheet.Id).OfCategory(BuiltInCategory.OST_Viewports).WhereElementIsNotElementType()
    for vp in vps:
        p = vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
        if p:
            s = p.AsString()
            if s:
                nums.add(s)
    return nums

def next_free_detail_number(preferred, taken):
    if not preferred:
        return None
    if preferred not in taken:
        return preferred
    try:
        n = int(preferred)
        while True:
            n += 1
            cand = str(n)
            if cand not in taken:
                return cand
    except:
        i = 1
        while True:
            cand = preferred + "-{}".format(i)
            if cand not in taken:
                return cand
            i += 1

def set_viewport_detail_number_safe(doc, sheet, vp, preferred):
    if not preferred:
        return
    taken = existing_detail_numbers_on_sheet(doc, sheet)
    pcur = vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
    if pcur:
        cur = pcur.AsString()
        if cur and cur in taken:
            taken.remove(cur)
    target = next_free_detail_number(preferred, taken)
    try:
        if target:
            p = vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
            if p and not p.IsReadOnly:
                p.Set(target)
            if target != preferred:
                logger.warning("    - Detail number '{}' taken on {}; set '{}' instead.".format(preferred, sheet.SheetNumber, target))
            else:
                logger.info("    - Restored detail number '{}' on {}.".format(target, sheet.SheetNumber))
    except Exception as e:
        logger.warning("    - Failed to set detail number on {}: {}".format(sheet.SheetNumber, e))

def sheet_has_view(doc, sheet, view_id):
    vps = FilteredElementCollector(doc, sheet.Id).OfCategory(BuiltInCategory.OST_Viewports).WhereElementIsNotElementType()
    for vp in vps:
        if vp.ViewId == view_id:
            return True
    return False

# -----------------------------
# Schedules (per-sheet scan with Revit 2024 method names)
# -----------------------------
def _sheet_schedule_instance_ids(doc, sheet):
    """Return list[ElementId] of ScheduleSheetInstance IDs on this sheet, trying known API shapes."""
    # 1) Revit 2024+: instance method
    try:
        ids = sheet.GetAllPlacedScheduleInstances()
        if ids: return list(ids)
    except:
        pass
    # 2) Some 2024 builds expose this name instead
    try:
        ids = sheet.GetSchedulePlacedInstances()
        if ids: return list(ids)
    except:
        pass
    # 3) Static helper on the class (older/common)
    try:
        ids = ScheduleSheetInstance.GetScheduleSheetInstances(doc, sheet.Id)
        if ids: return list(ids)
    except:
        pass
    # 4) Last resort: scoped collector
    try:
        return [ssi.Id for ssi in FilteredElementCollector(doc, sheet.Id).OfClass(ScheduleSheetInstance)]
    except:
        return []

def build_schedule_to_sheets_map(doc):
    """Return dict: schedule_view_id (int) -> list of {sheet_id, sheet_no, sheet_name, point}."""
    result = {}
    for sh in FilteredElementCollector(doc).OfClass(ViewSheet):
        for sid in _sheet_schedule_instance_ids(doc, sh):
            ssi = doc.GetElement(sid)
            if not isinstance(ssi, ScheduleSheetInstance):
                continue
            try:
                sched = doc.GetElement(ssi.ScheduleId)
                if not isinstance(sched, ViewSchedule):
                    continue
                key = sched.Id.IntegerValue
                lst = result.setdefault(key, [])
                pt = ssi.Point
                lst.append(dict(
                    sheet_id=sh.Id,
                    sheet_no=sh.SheetNumber,
                    sheet_name=sh.Name,
                    point=pt
                ))
            except:
                continue
    return result

def sheet_has_schedule(doc, sheet, schedule_id):
    """Check if sheet already has a ScheduleSheetInstance for schedule_id."""
    try:
        ssi_ids = _sheet_schedule_instance_ids(doc, sheet)
    except:
        ssi_ids = []
    for sid in ssi_ids:
        ssi = doc.GetElement(sid)
        try:
            if ssi.ScheduleId == schedule_id:
                return True
        except:
            pass
    return False

def place_schedule_on_sheet(doc, sheet, schedule_view, point_xyz):
    try:
        try:
            can_add = ScheduleSheetInstance.CanAddToSheet(doc, sheet.Id, schedule_view.Id)
            logger.info("    - CanAddToSheet({}, {}) = {}".format(sheet.SheetNumber, schedule_view.Name, can_add))
        except Exception as e:
            logger.info("    - CanAddToSheet check skipped/failed: {}".format(e))
        return ScheduleSheetInstance.Create(doc, sheet.Id, schedule_view.Id, point_xyz)
    except Exception as e:
        logger.error("    - EXCEPTION placing schedule '{}' on {}: {}".format(schedule_view.Name, sheet.SheetNumber, e))
        logger.error(traceback.format_exc())
        return None

def copy_schedule_itemize_flag(old_schedule, new_schedule):
    try:
        od, nd = old_schedule.Definition, new_schedule.Definition
        if hasattr(od, "IsItemized") and hasattr(nd, "IsItemized"):
            nd.IsItemized = od.IsItemized
            logger.info("    Copied 'Itemize every instance' = {}".format(od.IsItemized))
    except Exception as e:
        logger.warning("    - Could not copy 'Itemize every instance': {}".format(e))

# -----------------------------
# Discovery: find OLD/NEW pairs
# -----------------------------
def find_pairs(doc):
    all_views = list(FilteredElementCollector(doc).OfClass(View))
    logger.info("Views found (raw): {}".format(len(all_views)))
    name_to_views = {}
    for v in all_views:
        if not is_placeable_view(v):
            continue
        if is_schedule(v) and not ENABLE_SCHEDULES:
            continue
        name_to_views.setdefault(v.Name, []).append(v)

    pairs = []
    for name, vlist in name_to_views.items():
        base, num = strip_numeric_suffix(name)
        if num is None:
            continue
        old_list = name_to_views.get(base, [])
        if not old_list:
            continue
        for new_v in vlist:
            old_v = next((c for c in old_list if not c.IsTemplate), None)
            if old_v:
                pairs.append((old_v, new_v, base, num))
    pairs.sort(key=lambda t: (t[2], t[3]))
    return pairs[:1] if ONLY_FIRST_PAIR else pairs

# -----------------------------
# Logging helpers
# -----------------------------
def log_viewport_placements(placements):
    if not placements:
        logger.info("    No viewport placements found.")
        return
    logger.info("    Viewport placements ({}):".format(len(placements)))
    for rec in placements:
        c = rec["center"]
        logger.info("      - Sheet {} '{}': center=({:.3f},{:.3f}) typeId={} detail='{}'".format(
            rec["sheet_no"], rec["sheet_name"], c.X, c.Y,
            rec["vp_type_id"].IntegerValue if rec["vp_type_id"] else None,
            rec["detail_number"] if rec["detail_number"] else ""
        ))

def log_schedule_placements(label, placements):
    if not placements:
        logger.info("    {}: none".format(label)); return
    logger.info("    {} ({}):".format(label, len(placements)))
    for rec in placements:
        p = rec["point"]
        logger.info("      - Sheet {} '{}': point=({:.3f},{:.3f})".format(
            rec["sheet_no"], rec["sheet_name"], p.X, p.Y))

def sheet_label(doc, sheet_id):
    sh = doc.GetElement(sheet_id)
    return u"{} '{}'".format(sh.SheetNumber, sh.Name) if isinstance(sh, ViewSheet) else u"id={}".format(sheet_id.IntegerValue)

# -----------------------------
# Processing
# -----------------------------
def process_pair(doc, old_view, new_view, base_name, sched_map):
    is_sched = is_schedule(old_view)

    # Capture placements for this pair
    if is_sched:
        placements_sc = list(sched_map.get(old_view.Id.IntegerValue, []))
        # if NEW already placed somewhere (rare), union those too
        if new_view.Id.IntegerValue in sched_map:
            seen = set((r["sheet_id"].IntegerValue for r in placements_sc))
            for r in sched_map[new_view.Id.IntegerValue]:
                if r["sheet_id"].IntegerValue not in seen:
                    placements_sc.append(r)
        placements_vp = []
    else:
        placements_vp = collect_viewport_placements(doc, old_view.Id)
        placements_sc = []

    tos_value = get_title_on_sheet(old_view)

    logger.info("  > Preparing swap for base '{}' ({})".format(base_name, "Schedule" if is_sched else "View/Legend"))
    logger.info("    OLD: '{}' id={}".format(old_view.Name, old_view.Id.IntegerValue))
    logger.info("    NEW: '{}' id={}".format(new_view.Name, new_view.Id.IntegerValue))
    logger.info("    Title on Sheet (OLD): '{}'".format(tos_value if tos_value else "<none>"))

    if is_sched:
        log_schedule_placements("Placements (per-sheet scan)", placements_sc)
    else:
        log_viewport_placements(placements_vp)

    # If schedules still have no placements, protect the sheet: archive OLD; do not delete.
    if is_sched and SAFE_SKIP_DELETE_IF_NO_PLACEMENTS and len(placements_sc) == 0:
        logger.warning("    No schedule placements detected for '{}' — archiving OLD; NOT deleting.".format(base_name))
        if not DRY_RUN:
            t = Transaction(doc, "Rename schedules (no-delete safety)")
            t.Start()
            try:
                old_tmp_name = base_name + "__OLD_ARCHIVE"
                try:
                    logger.info("    Archiving OLD '{}' → '{}'".format(old_view.Name, old_tmp_name))
                    old_view.Name = old_tmp_name
                except Exception as e:
                    logger.warning("    - Archive rename failed: {}".format(e))
                try:
                    logger.info("    Renaming NEW '{}' → '{}'".format(new_view.Name, base_name))
                    new_view.Name = base_name
                except Exception as e2:
                    logger.warning("    - Rename NEW to base failed: {}".format(e2))
                set_title_on_sheet(new_view, tos_value)
                copy_schedule_itemize_flag(old_view, new_view)
                t.Commit()
            except:
                try: t.RollBack()
                except: pass
        return True

    if DRY_RUN:
        logger.info("    [DRY RUN] Would rename, place, then delete OLD_TMP.")
        return True

    # Normal path
    t = Transaction(doc, "Safe Swap '{}' ⇢ '{}'".format(old_view.Name, new_view.Name))
    t.Start()
    try:
        # 1) Rename to prevent name collisions while placed
        old_tmp_name = base_name + "__OLD_TMP"
        try:
            logger.info("    Renaming OLD '{}' → '{}'".format(old_view.Name, old_tmp_name))
            old_view.Name = old_tmp_name
        except Exception as e:
            logger.warning("    - Could not rename OLD to temp: {}".format(e))
        try:
            logger.info("    Renaming NEW '{}' → '{}'".format(new_view.Name, base_name))
            new_view.Name = base_name
        except Exception as e_rename:
            fb = base_name + " (relinked)"
            logger.warning("    - Rename NEW failed ({}). Using fallback '{}'".format(e_rename, fb))
            new_view.Name = fb

        # 2) Restore parameters onto NEW before placement
        set_title_on_sheet(new_view, tos_value)
        if is_sched:
            copy_schedule_itemize_flag(old_view, new_view)

        # 3) Place NEW instances
        created_vps, created_ssi = [], []
        if is_sched:
            target = len(placements_sc)
            for rec in placements_sc:
                sheet = doc.GetElement(rec["sheet_id"])
                if not isinstance(sheet, ViewSheet):
                    logger.warning("    - Skipping non-sheet id={}".format(rec["sheet_id"].IntegerValue)); continue
                logger.info("    - TARGET SHEET: {} | point=({:.3f},{:.3f})".format(
                    sheet_label(doc, rec["sheet_id"]), rec["point"].X, rec["point"].Y))
                if sheet_has_schedule(doc, sheet, new_view.Id):
                    logger.info("    - Sheet already has this schedule; skipping."); continue
                inst = place_schedule_on_sheet(doc, sheet, new_view, rec["point"])
                if inst: created_ssi.append((inst.Id, sheet.Id))
                else: logger.error("    - FAILED creating schedule on {}".format(sheet_label(doc, rec["sheet_id"])))
            if len(created_ssi) < target:
                logger.error("    VERIFICATION FAILED: placed {} of {} schedule instance(s). Rolling back."
                             .format(len(created_ssi), target))
                t.RollBack(); return False
        else:
            for rec in placements_vp:
                sheet = doc.GetElement(rec["sheet_id"])
                if not isinstance(sheet, ViewSheet):
                    logger.warning("    - Skipping non-sheet id={}".format(rec["sheet_id"].IntegerValue)); continue
                logger.info("    - Placing view on {} at center=({:.3f},{:.3f})".format(
                    sheet_label(doc, rec["sheet_id"]), rec["center"].X, rec["center"].Y))
                if sheet_has_view(doc, sheet, new_view.Id):
                    logger.info("    - Sheet already has this view; skipping."); continue
                vp = place_view_on_sheet(doc, sheet, new_view, rec["center"], rec["vp_type_id"])
                if vp: created_vps.append((sheet, vp, rec["detail_number"]))
                else: logger.error("    - FAILED creating viewport on {}".format(sheet_label(doc, rec["sheet_id"])))

        # 4) Delete OLD temp
        try:
            logger.info("    Deleting OLD temp view '{}'".format(old_tmp_name))
            doc.Delete(old_view.Id)
        except Exception as e_del:
            logger.warning("    - Could not delete OLD temp: {}".format(e_del))

        # 5) Views: restore detail numbers after OLD is gone
        for (sheet, vp, preferred_dn) in created_vps:
            set_viewport_detail_number_safe(doc, sheet, vp, preferred_dn)

        # 6) Post-check (optional)
        if is_sched:
            finals = build_schedule_to_sheets_map(doc).get(new_view.Id.IntegerValue, [])
            logger.info("    POST-CHECK: '{}' now placed on {} sheet(s).".format(new_view.Name, len(finals)))
            for rec in finals:
                logger.info("      - {}".format(sheet_label(doc, rec["sheet_id"])))

        t.Commit()
        logger.info("    Swap complete for base '{}'.".format(base_name))
        return True

    except Exception as e_txn:
        logger.error("    ERROR during swap for base '{}': {}".format(base_name, e_txn))
        logger.error(traceback.format_exc())
        try: t.RollBack()
        except: pass
        return False

# -----------------------------
# Main
# -----------------------------
def main():
    logger.info("---- start ----")
    doc = __revit__.ActiveUIDocument.Document

    # Build a definitive schedule→sheets map by scanning every sheet
    sched_map = build_schedule_to_sheets_map(doc)
    if EXTRA_DIAG:
        total_links = sum(len(v) for v in sched_map.values())
        logger.info("Schedule map: {} schedules placed across {} sheet links."
                    .format(len(sched_map), total_links))

    pairs = find_pairs(doc)
    if not pairs:
        logger.warning("No (OLD base name, NEW with numeric suffix) pairs found.")
        output.print_md("> **No pairs found.** Example: `Door Schedule` + `Door Schedule1`, or `L1 - Plan` + `L1 - Plan1`.")
        return

    logger.info("Pairs to process: {}".format(len(pairs)))
    for (old_view, new_view, base_name, suffix_num) in pairs:
        logger.info("  - Pair: OLD='{}'  NEW='{}'  -> base='{}'  num={}  type={}".format(
            old_view.Name, new_view.Name, base_name, suffix_num,
            "Schedule" if is_schedule(old_view) else "View/Legend"
        ))

    affected = 0
    tg = TransactionGroup(doc, "Safe Replace Views & Schedules (per-sheet scan)")
    tg.Start()
    try:
        for (old_view, new_view, base_name, _num) in pairs:
            if old_view.Id == new_view.Id:
                logger.warning("  ! Skipping pair where OLD and NEW are same element: '{}'".format(old_view.Name))
                continue
            if process_pair(doc, old_view, new_view, base_name, sched_map):
                affected += 1

        if affected > 0:
            tg.Assimilate()
            logger.info("---- done; processed {} pair(s) ----".format(affected))
        else:
            tg.RollBack()
            logger.warning("Nothing was processed; rolled back TransactionGroup.")
    except Exception as e_fatal:
        logger.error("FATAL: {}".format(e_fatal))
        logger.error(traceback.format_exc())
        try: tg.RollBack()
        except: pass

if __name__ == "__main__":
    main()
