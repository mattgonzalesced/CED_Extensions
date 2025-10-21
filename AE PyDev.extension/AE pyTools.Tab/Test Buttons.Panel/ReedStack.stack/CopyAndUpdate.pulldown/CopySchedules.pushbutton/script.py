# -*- coding: utf-8 -*-
# Copy selected schedules (standard + panel) to other open documents
# - No placements
# - No renaming (let Revit auto-suffix in destination)
# - One Transaction PER TARGET DOCUMENT (batch copies all selected schedules to that target)
# - Prevents newly opened schedule tabs in the ACTIVE document (restores your active view; closes new schedule UIViews)
#
# Revit 2024 / IronPython 2.7

from __future__ import print_function
import traceback

from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSchedule, ElementId,
    ElementTransformUtils, CopyPasteOptions, IDuplicateTypeNamesHandler,
    Transaction, Transform
)

# Panel schedules
try:
    from Autodesk.Revit.DB.Electrical import PanelScheduleView
except:
    try:
        from Autodesk.Revit.DB import PanelScheduleView
    except:
        PanelScheduleView = None

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

def is_placeable_panelschedule(psv):
    if PanelScheduleView is None: return False
    try:
        if hasattr(psv, "IsTemplate") and psv.IsTemplate: return False
    except: pass
    return isinstance(psv, PanelScheduleView)

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


# ---------- UI selection ----------
def pick_schedules(source_doc):
    vs_list = [v for v in FilteredElementCollector(source_doc).OfClass(ViewSchedule) if is_placeable_viewschedule(v)]
    ps_list = []
    if PanelScheduleView is not None:
        try:
            ps_list = [p for p in FilteredElementCollector(source_doc).OfClass(PanelScheduleView) if is_placeable_panelschedule(p)]
        except: ps_list = []

    if not vs_list and not ps_list:
        forms.alert("No copyable schedules (standard or panel) found in the active document.", exitscript=True)

    items = []
    for v in vs_list:
        items.append(ListItem(u"[Schedule]  {}  (id:{})".format(v.Name, v.Id.IntegerValue), ("VS", v)))
    for p in ps_list:
        items.append(ListItem(u"[Panel]     {}  (id:{})".format(p.Name, p.Id.IntegerValue), ("PS", p)))

    selected = forms.SelectFromList.show(items, title="Select schedules to copy",
                                         multiselect=True, width=900, height=650, button_name="Copy")
    if not selected: return []
    return [itm.data for itm in selected]   # list of ("VS"/"PS", element)

def ask_targets(target_docs):
    if not target_docs:
        forms.alert("No other open Revit documents detected.", exitscript=True)
    items = [ListItem("{}{}".format(td.Title, "" if td.IsWorkshared else " (non-workshared)"), td, checked=True) for td in target_docs]
    selected = forms.SelectFromList.show(items, title="Copy to which open documents?",
                                         multiselect=True, width=700, height=500, button_name="Use These")
    if not selected: return []
    return [itm.data for itm in selected]


# ---------- Copy helpers (batch per target) ----------
class _DupTypeUseDestination(IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        # Keep destination types on name collisions
        return args.UseDestinationTypes()

def batch_copy_to_target(src_doc, target_doc, element_ids):
    """
    Copy a batch of elements from src → target inside a single Transaction on the target.
    element_ids: iterable of ElementId (from the source doc)
    Returns mapped ids or None on failure.
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
        t.Commit()
        return mapped
    except Exception as e:
        logger.error("CopyElements failed into '{}': {}".format(target_doc.Title, e))
        logger.error(traceback.format_exc())
        try: t.RollBack()
        except: pass
        return None


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
                # Close brand-new schedule tabs
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


# ---------- Main ----------
def main():
    picked = pick_schedules(doc)   # list of ("VS"/"PS", element)
    if not picked: return

    targets = ask_targets(get_open_target_docs(doc))
    if not targets: return

    # Gather all source element ids to copy (we copy both types together)
    src_ids = []
    for kind, elem in picked:
        src_ids.append(elem.Id)

    total_vs = sum(1 for k,_ in picked if k=="VS")
    total_ps = sum(1 for k,_ in picked if k=="PS")

    output.print_md("### Copying schedules")
    output.print_md("- Standard schedules selected: **{}**".format(total_vs))
    output.print_md("- Panel schedules selected: **{}**".format(total_ps))
    output.print_md("- Target docs: **{}**".format(len(targets)))
    output.print_md("- No placements. No renaming. Revit will auto-suffix on conflicts.")
    output.print_md("- One transaction **per target document** (API limitation).")

    # Snapshot active UI state to avoid schedule tabs popping open
    original_active_view = uidoc.ActiveView
    initially_open_ids = get_active_doc_open_viewids()

    for td in targets:
        logger.info("=== Target: {} ===".format(td.Title))

        # Batch copy (one transaction) for ALL selected schedules into this target
        mapped = batch_copy_to_target(doc, td, src_ids)
        if mapped is None:
            output.print_md("- **{}**: copy failed".format(td.Title))
            continue

        # Count what we intended to copy (just for summary)
        output.print_md("- **{}**: copied **{}** schedules (std + panel)".format(td.Title, len(src_ids)))

    # UI cleanup in the active doc
    close_new_schedule_tabs_in_active_doc(initially_open_ids)
    restore_active_view(original_active_view)

    output.print_md("### ✅ Done")
    output.print_md("- Copies created in each target doc inside a single transaction per target.")
    output.print_md("- Active view restored; any newly opened schedule tabs were closed.")

if __name__ == "__main__":
    main()
