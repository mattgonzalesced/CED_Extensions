# -*- coding: utf-8 -*-
# Copy selected schedules (STANDARD ONLY) to other open documents
# - No renaming during copy (let Revit auto-suffix in destination)
# - One Transaction PER TARGET DOCUMENT (batch copies all selected schedules to that target)
# - Prevents newly opened schedule tabs in the ACTIVE document (restores your active view; closes new schedule UIViews)
#
# Replacement flow (standard schedules only):
#   • After copy, if a new schedule name is auto-suffixed (e.g., "Name 1"):
#       - Find the existing schedule in the target named "Name"
#       - Record all of THAT schedule's placements on sheets (sheet id, point, rotation)
#       - Place the new schedule in the SAME locations on the SAME sheets
#       - Delete the old schedule
#       - Rename the new schedule to the original name ("Name")
#
# Revit 2024 / IronPython 2.7

from __future__ import print_function

import re
import traceback

from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSchedule, ElementId,
    ElementTransformUtils, CopyPasteOptions, IDuplicateTypeNamesHandler,
    Transaction, Transform, XYZ
)

# Sheet schedule instances (standard schedules)
try:
    from Autodesk.Revit.DB import ScheduleSheetInstance
except:
    ScheduleSheetInstance = None

from System.Collections.Generic import List  # for ICollection[ElementId]
from pyrevit import script, forms

uiapp = __revit__
uidoc = uiapp.ActiveUIDocument
doc   = uidoc.Document
app   = uiapp.Application

logger = script.get_logger()
output = script.get_output()


# ---------- Small helper so SelectFromList works everywhere ----------
class ListItem(object):
    def __init__(self, label, data, checked=False):
        self.label = label
        self.data = data
        self.checked = checked
    def __str__(self): return self.label


# ---------- Utilities ----------
def is_placeable_viewschedule(vs):
    try:
        if vs.IsTemplate: return False
    except: pass
    return isinstance(vs, ViewSchedule)

def get_open_target_docs(source_doc):
    targets = []
    for d in app.Documents:
        try:
            if d.IsLinked: continue
        except: pass
        if d.IsFamilyDocument: continue
        # skip the source
        try:
            if d.PathName == source_doc.PathName and d.Title == source_doc.Title:
                continue
        except:
            if d.Title == source_doc.Title:
                continue
        targets.append(d)
    return targets


# ---------- UI selection (STANDARD schedules only) ----------
def pick_schedules(source_doc):
    vs_list = [v for v in FilteredElementCollector(source_doc).OfClass(ViewSchedule) if is_placeable_viewschedule(v)]
    if not vs_list:
        forms.alert("No copyable standard schedules found in the active document.", exitscript=True)

    items = []
    for v in vs_list:
        items.append(ListItem(u"[Schedule]  {}  (id:{})".format(v.Name, v.Id.IntegerValue), ("VS", v)))

    selected = forms.SelectFromList.show(items, title="Select standard schedules to copy",
                                         multiselect=True, width=900, height=650, button_name="Copy")
    if not selected: return []
    return [itm.data for itm in selected]   # list of ("VS", element)

def ask_targets(target_docs):
    if not target_docs:
        forms.alert("No other open Revit documents detected.", exitscript=True)
    items = [ListItem("{}{}".format(td.Title, "" if td.IsWorkshared else " (non-workshared)"), td, checked=True) for td in target_docs]
    selected = forms.SelectFromList.show(items, title="Copy to which open documents?",
                                         multiselect=True, width=700, height=500, button_name="Use These")
    if not selected: return []
    return [itm.data for itm in selected]


# ---------- Copy helpers (batch per target) ----------
from Autodesk.Revit.DB import DuplicateTypeAction

class _DupTypeUseDestination(IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        try:
            args.SetAction(DuplicateTypeAction.UseDestinationTypes)
        except:
            # Compatibility with pre-2024 Revit versions
            try:
                return args.UseDestinationTypes()
            except:
                pass
        return


def batch_copy_to_target(src_doc, target_doc, element_ids):
    """
    Copy a batch of elements from src → target inside a single Transaction on the target.
    element_ids: iterable of ElementId (from the source doc)
    Returns (Transaction, ICollection[ElementId]) for mapped new ids, or (None, None) on failure.
    """
    ids = List[ElementId]()
    for eid in element_ids:
        ids.Add(eid)

    opts = CopyPasteOptions()
    opts.SetDuplicateTypeNamesHandler(_DupTypeUseDestination())

    t = Transaction(target_doc, "Copy schedules from '{}'".format(src_doc.Title))
    t.Start()
    try:
        mapped = ElementTransformUtils.CopyElements(src_doc, ids, target_doc, Transform.Identity, opts)
        # Replacement runs inside this same transaction for atomicity.
        return (t, mapped)
    except Exception as e:
        logger.error("CopyElements failed into '{}': {}".format(target_doc.Title, e))
        logger.error(traceback.format_exc())
        try: t.RollBack()
        except: pass
        return (None, None)


# ---------- Keep UI from opening schedule tabs ----------
def get_active_doc_open_viewids():
    """Return set of view ElementId.IntegerValue currently open as UIViews in the ACTIVE document."""
    try:
        return set([uiv.ViewId.IntegerValue for uiv in uidoc.GetOpenUIViews()])
    except:
        return set()

def close_new_schedule_tabs_in_active_doc(initial_open_ids):
    """Close any newly-opened schedule UIViews in the ACTIVE document."""
    try:
        now_views = list(uidoc.GetOpenUIViews())
    except:
        return
    for uiv in now_views:
        vid = uiv.ViewId.IntegerValue
        if vid in initial_open_ids:
            continue
        try:
            v = doc.GetElement(uiv.ViewId)
            if isinstance(v, ViewSchedule):
                try:
                    uiv.Close()
                except:
                    pass
        except:
            pass

def restore_active_view(original_view):
    """Switch back to the original active view (prevents Revit from leaving a schedule active)."""
    try:
        if original_view and original_view.Id != uidoc.ActiveView.Id:
            uidoc.RequestViewChange(original_view)
    except:
        pass


# ---------- Sheet placement helpers (standard schedules) ----------
_SUFFIX_RE = re.compile(r"^(.*?)[\s]*(\d+)$")  # catches "... 1", "... 2", also "...1"

def _basename_if_suffixed(name):
    """If Revit appended a numeric suffix (e.g., 'Name 1'), return ('Name', number). Else (name, None)."""
    m = _SUFFIX_RE.match(name)
    if not m:
        return (name, None)
    base = m.group(1).rstrip()
    num  = m.group(2)
    return (base, num)

def _collect_old_vs_by_name(target_doc):
    """Return dict name -> ViewSchedule (non-template)."""
    d = {}
    for vs in FilteredElementCollector(target_doc).OfClass(ViewSchedule):
        if is_placeable_viewschedule(vs):
            try:
                d[vs.Name] = vs
            except:
                pass
    return d

def _gather_schedule_sheet_instances_std(target_doc, vs):
    """Return a list of dicts with placement for a standard schedule on sheets."""
    placements = []
    if ScheduleSheetInstance is None:
        return placements
    try:
        ssi_col = FilteredElementCollector(target_doc).OfClass(ScheduleSheetInstance)
        for ssi in ssi_col:
            try:
                if ssi.ScheduleId == vs.Id:
                    sheet_id = ssi.OwnerViewId  # the sheet view id
                    pt = None
                    rot = None
                    try:
                        pt = ssi.Point  # XYZ in sheet coordinates
                    except:
                        pt = XYZ(0,0,0)
                    try:
                        rot = ssi.Rotation  # ScheduleRotation enum
                    except:
                        rot = None
                    placements.append({"sheet_id": sheet_id, "point": pt, "rotation": rot})
            except:
                continue
    except:
        pass
    return placements

def _place_schedule_instances_std(target_doc, new_vs, placements):
    """Create instances of a standard schedule on given sheets with same point/rotation."""
    created = []
    if ScheduleSheetInstance is None:
        return created
    for p in placements:
        try:
            ssi = ScheduleSheetInstance.Create(target_doc, p["sheet_id"], new_vs.Id, p["point"])
            try:
                if p.get("rotation") is not None:
                    ssi.Rotation = p["rotation"]
            except:
                pass
            created.append(ssi)
        except Exception as e:
            logger.warning("Failed placing new schedule '{}' on sheet {}: {}".format(new_vs.Name, p["sheet_id"].IntegerValue, e))
    return created


# ---------- Replacement workflow (STANDARD ONLY) ----------
def _post_copy_replace_in_target(target_doc, picked_pairs, src_ids, mapped_new_ids):
    """
    For each copied standard schedule:
      If new name is auto-suffixed (e.g., 'X 1'), treat 'X' in target as the OLD schedule.
      Replicate sheet placements from OLD -> NEW, delete OLD, rename NEW to base name.
    """
    # Build name->ViewSchedule dict (target state before deletions)
    vs_by_name = _collect_old_vs_by_name(target_doc)

    # Map src -> new elem
    new_by_src = {}  # src ElementId -> new element
    for i, src_eid in enumerate(src_ids):
        try:
            new_id = list(mapped_new_ids)[i]
            new_elem = target_doc.GetElement(new_id)
            new_by_src[src_eid.IntegerValue] = new_elem
        except:
            continue

    replaced_count = 0

    for (kind, src_elem) in picked_pairs:
        if kind != "VS":
            continue

        new_elem = new_by_src.get(src_elem.Id.IntegerValue, None)
        if new_elem is None or not isinstance(new_elem, ViewSchedule):
            continue

        try:
            new_name = new_elem.Name
        except:
            continue

        base_name, suffix_num = _basename_if_suffixed(new_name)
        if suffix_num is None:
            # No conflict, nothing to replace
            continue

        try:
            old_vs = vs_by_name.get(base_name, None)
            if old_vs is None:
                continue

            # Gather placements from OLD
            placements = _gather_schedule_sheet_instances_std(target_doc, old_vs)

            # Place NEW instances
            _place_schedule_instances_std(target_doc, new_elem, placements)

            # Delete OLD then rename NEW to base name
            try:
                target_doc.Delete(old_vs.Id)
            except Exception as e:
                logger.warning("Failed to delete old schedule '{}': {}".format(old_vs.Name, e))
            try:
                new_elem.Name = base_name
            except Exception as e:
                logger.warning("Failed to rename '{}' to '{}': {}".format(new_name, base_name, e))

            replaced_count += 1

        except Exception as ex:
            logger.error("Replacement flow failed for '{}': {}".format(new_name, ex))
            logger.error(traceback.format_exc())

    return replaced_count


# ---------- Main ----------
def main():
    picked = pick_schedules(doc)   # list of ("VS", element)
    if not picked: return

    targets = ask_targets(get_open_target_docs(doc))
    if not targets: return

    # Gather all source element ids to copy
    src_ids = [elem.Id for kind, elem in picked if kind == "VS"]

    total_vs = len(src_ids)

    output.print_md("### Copying standard schedules")
    output.print_md("- Standard schedules selected: **{}**".format(total_vs))
    output.print_md("- Target docs: **{}**".format(len(targets)))
    output.print_md("- No renaming during copy; Revit may auto-suffix on conflicts.")
    output.print_md("- One transaction **per target document** (API limitation).")
    output.print_md("- On conflict, placements are replicated from the original, the original deleted, and the new copy renamed to the original name.")

    # Snapshot active UI state to avoid schedule tabs popping open
    original_active_view = uidoc.ActiveView
    initially_open_ids = get_active_doc_open_viewids()

    for td in targets:
        logger.info("=== Target: {} ===".format(td.Title))

        # Batch copy (one transaction) for ALL selected schedules into this target
        t, mapped = batch_copy_to_target(doc, td, src_ids)
        if t is None or mapped is None:
            output.print_md("- **{}**: copy failed".format(td.Title))
            continue

        # Replacement workflow runs inside the SAME transaction for atomicity
        try:
            replaced = _post_copy_replace_in_target(td, picked, src_ids, mapped)
            t.Commit()
        except Exception as e:
            logger.error("Replacement flow error for target '{}': {}".format(td.Title, e))
            logger.error(traceback.format_exc())
            try: t.RollBack()
            except: pass
            output.print_md("- **{}**: copy/replace failed (rolled back)".format(td.Title))
            continue

        output.print_md("- **{}**: copied **{}** schedule(s); replaced (auto-suffix conflicts): **{}**"
                        .format(td.Title, len(src_ids), replaced))

    # UI cleanup in the active doc
    close_new_schedule_tabs_in_active_doc(initially_open_ids)
    restore_active_view(original_active_view)

    output.print_md("### ✅ Done")
    output.print_md("- Copies created in each target doc inside a single transaction per target.")
    output.print_md("- For name conflicts, placements were replicated from the original, the original deleted, and the new copy was renamed.")
    output.print_md("- Active view restored; any newly opened schedule tabs were closed.")

if __name__ == "__main__":
    main()

