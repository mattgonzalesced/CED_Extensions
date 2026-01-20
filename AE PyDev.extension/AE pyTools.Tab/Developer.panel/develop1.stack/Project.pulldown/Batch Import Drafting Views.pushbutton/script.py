# -*- coding: utf-8 -*-
import os
import sys

from pyrevit import revit, forms, DB, script
from pyrevit.framework import List

logger = script.get_logger()
output = script.get_output()

VIEW_TOS_PARAM = DB.BuiltInParameter.VIEW_DESCRIPTION


# ------------------------------------------------------------
# UI
# ------------------------------------------------------------
def pick_root_folder():
    folder = forms.pick_folder("Select folder containing RVT files")
    if not folder:
        sys.exit()
    return folder


def find_rvt_files(folder):
    rvts = []
    for root, dirs, files in os.walk(folder):
        for f in files:
            if f.lower().endswith(".rvt"):
                rvts.append(os.path.join(root, f))
    return rvts


def pick_files(files):
    names = [os.path.basename(f) for f in files]
    selected = forms.SelectFromList.show(
        names,
        multiselect=True,
        title="Select RVT files",
        button_name="Import"
    )
    if not selected:
        sys.exit()

    out = []
    for name in selected:
        for fp in files:
            if fp.endswith(name):
                out.append(fp)
                break
    return out


# ------------------------------------------------------------
# COPY ENGINE
# ------------------------------------------------------------
class CopyUseDestination(DB.IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DB.DuplicateTypeAction.UseDestinationTypes


def get_detail_elements(doc, view):
    coll = DB.FilteredElementCollector(doc, view.Id)
    coll.WhereElementIsNotElementType()
    result = []

    for el in coll:
        # Skip any actual view objects
        if isinstance(el, DB.View):
            continue

        # Skip viewports and extent elems
        name = ""
        try:
            name = el.Name or ""
        except:
            pass

        if isinstance(el, DB.Viewport):
            continue

        if "extent" in name.lower():
            continue

        # Skip guide grids
        if el.Category and "guide" in el.Category.Name.lower():
            continue

        # Skip any category called "Views"
        if el.Category and el.Category.Name.lower() == "views":
            continue

        # Skip sheet references or other sheet-view links
        if el.Category and "sheet" in el.Category.Name.lower():
            continue

        result.append(el.Id)

    return result


def create_target_view(dest_doc, source_view):
    """Create a new empty drafting view with matching name + settings."""
    vtype = dest_doc.GetDefaultElementTypeId(DB.ElementTypeGroup.ViewTypeDrafting)

    with revit.Transaction("Create Drafting View", doc=dest_doc):
        newv = DB.ViewDrafting.Create(dest_doc, vtype)
        try:
            newv.Name = source_view.Name
        except:
            pass

        # Copy scale
        try:
            newv.Scale = source_view.Scale
        except:
            pass

        # Copy Title On Sheet
        try:
            sp = source_view.Parameter[VIEW_TOS_PARAM]
            dp = newv.Parameter[VIEW_TOS_PARAM]
            if sp and dp and not dp.IsReadOnly:
                dp.Set(sp.AsString())
        except:
            pass

    return newv


def copy_drafting_contents(src_doc, src_view, dest_doc, dest_view):
    """Copy ONLY the detailing from src → dest."""
    ids = get_detail_elements(src_doc, src_view)
    if not ids:
        print("  Skipped empty view:", src_view.Name)
        return

    cp = DB.CopyPasteOptions()
    cp.SetDuplicateTypeNamesHandler(CopyUseDestination())

    with revit.Transaction("Copy Drafting Contents", doc=dest_doc, swallow_errors=True):
        DB.ElementTransformUtils.CopyElements(
            src_view,
            List[DB.ElementId](ids),
            dest_view,
            None,
            cp
        )

def find_dest_view(dest_doc, name):
    """Find a drafting view in the destination doc by exact name."""
    col = DB.FilteredElementCollector(dest_doc) \
              .OfClass(DB.ViewDrafting) \
              .WhereElementIsNotElementType()

    for v in col:
        if not v.IsTemplate and v.Name == name:
            return v

    return None

def copy_single_drafting_view(src_doc, src_view, dest_doc):
    name = src_view.Name

    # Find destination view with same name (if any)
    existing = find_dest_view(dest_doc, name)

    with revit.Transaction("Create/Prepare Drafting View", doc=dest_doc):
        # Create or reuse
        if existing:
            target = existing
            # Clear existing contents only (not deleting view)
            ids = get_detail_elements(dest_doc, target)
            for eid in ids:
                try:
                    dest_doc.Delete(eid)
                except:
                    pass
        else:
            # Create new drafting view
            vtype = dest_doc.GetDefaultElementTypeId(DB.ElementTypeGroup.ViewTypeDrafting)
            target = DB.ViewDrafting.Create(dest_doc, vtype)

            try:
                target.Name = name
            except:
                pass

        # Copy basic view properties
        try:
            target.Scale = src_view.Scale
        except:
            pass

        try:
            sp = src_view.Parameter[VIEW_TOS_PARAM]
            dp = target.Parameter[VIEW_TOS_PARAM]
            if sp and dp and not dp.IsReadOnly:
                dp.Set(sp.AsString())
        except:
            pass

    # --------- Copy elements into the view ---------
    ids = get_detail_elements(src_doc, src_view)
    if not ids:
        print("  Skipped (empty view):", name)
        return

    cp = DB.CopyPasteOptions()
    cp.SetDuplicateTypeNamesHandler(CopyUseDestination())

    with revit.Transaction("Copy Drafting Contents", doc=dest_doc, swallow_errors=True):
        DB.ElementTransformUtils.CopyElements(
            src_view,
            List[DB.ElementId](ids),
            target,
            None,
            cp
        )

    print("  Imported:", name)



# ------------------------------------------------------------
# IMPORT RVT FILE
# ------------------------------------------------------------
def import_from_rvt(app, dest_doc, path):
    print("\nSOURCE:", os.path.basename(path))

    # Open source doc without activating it
    try:
        mpath = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(path)
        opts = DB.OpenOptions()
        opts.Audit = False
        opts.DetachFromCentralOption = DB.DetachFromCentralOption.DetachAndPreserveWorksets
        opts.AllowOpeningLocalByWrongUser = True

        src_doc = app.OpenDocumentFile(mpath, opts)

    except Exception as err:
        print("  ERROR opening:", err)
        return

    try:
        drafts = DB.FilteredElementCollector(src_doc)\
            .OfClass(DB.ViewDrafting)\
            .WhereElementIsNotElementType()\
            .ToElements()

        drafts = [v for v in drafts if not v.IsTemplate]
        print("  Drafting views found:", len(drafts))

        for dv in drafts:
            copy_single_drafting_view(src_doc, dv, dest_doc)

    finally:
        try:
            src_doc.Close(False)
        except:
            pass


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------
def main():
    dest_doc = revit.doc
    app = dest_doc.Application

    folder = pick_root_folder()
    files = find_rvt_files(folder)

    if not files:
        print("No RVT files found.")
        return

    selected = pick_files(files)

    tg = DB.TransactionGroup(dest_doc, "Batch Import Drafting Views")
    tg.Start()

    try:
        for fp in selected:
            import_from_rvt(app, dest_doc, fp)

        tg.Assimilate()
    except:
        tg.RollBack()
        raise

    print("\nDone.\n")


main()
