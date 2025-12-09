# -*- coding: utf-8 -*-
import os
import sys

from pyrevit import revit, forms, DB, script
from pyrevit.framework import List

logger = script.get_logger()
output = script.get_output()

VIEW_TOS_PARAM = DB.BuiltInParameter.VIEW_DESCRIPTION
OPTIONS_XAML_PATH = script.get_bundle_file('import_drafting_options.xaml')


# ------------------------------------------------------------
# OPTIONS WINDOW (WPF)
# ------------------------------------------------------------
class DraftingImportOptionsWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, OPTIONS_XAML_PATH)
        self.mode = 'multi'            # 'single' or 'multi'
        self.existing_policy = 'update'  # 'skip', 'update', 'duplicate'
        self._help_key = None

        self._bind_events()
        self._init_defaults()
        self._set_help('mode_multi')

    def _bind_events(self):
        # Mode radios
        self.single_mode_rb.Checked += self._on_mode_changed
        self.multi_mode_rb.Checked += self._on_mode_changed
        self.single_mode_rb.GotFocus += lambda s, e: self._set_help('mode_single')
        self.multi_mode_rb.GotFocus += lambda s, e: self._set_help('mode_multi')

        # Existing behavior radios
        self.skip_existing_rb.Checked += self._on_existing_changed
        self.update_existing_rb.Checked += self._on_existing_changed
        self.duplicate_existing_rb.Checked += self._on_existing_changed

        self.skip_existing_rb.GotFocus += lambda s, e: self._set_help('existing_skip')
        self.update_existing_rb.GotFocus += lambda s, e: self._set_help('existing_update')
        self.duplicate_existing_rb.GotFocus += lambda s, e: self._set_help('existing_duplicate')

        # Buttons
        self.ok_btn.Click += self._on_ok
        self.cancel_btn.Click += self._on_cancel

    def _init_defaults(self):
        # Defaults already set in XAML (multi + update)
        self.mode = 'multi'
        self.existing_policy = 'update'

    def _on_mode_changed(self, sender, args):
        tag = getattr(sender, 'Tag', None)
        if tag in ('single', 'multi'):
            self.mode = tag
        if tag == 'single':
            self._set_help('mode_single')
        elif tag == 'multi':
            self._set_help('mode_multi')

    def _on_existing_changed(self, sender, args):
        tag = getattr(sender, 'Tag', None)
        if tag in ('skip', 'update', 'duplicate'):
            self.existing_policy = tag
        if tag == 'skip':
            self._set_help('existing_skip')
        elif tag == 'update':
            self._set_help('existing_update')
        elif tag == 'duplicate':
            self._set_help('existing_duplicate')

    def _on_ok(self, sender, args):
        # Just close; main() will read self.mode / self.existing_policy
        self.Close()

    def _on_cancel(self, sender, args):
        # Abort whole script
        script.exit()

    def _help_texts(self):
        return {
            'mode_single': (
                "Insert from single file:\n"
                "- Pick one RVT file.\n"
                "- Then choose which drafting views from that file to import."
            ),
            'mode_multi': (
                "Insert from multiple files:\n"
                "- Pick a root folder.\n"
                "- The tool finds RVT files.\n"
                "- You choose which RVT files to import.\n"
                "- All drafting views from those files are brought in."
            ),
            'existing_skip': (
                "Skip existing views:\n"
                "- If a drafting view with the same name already exists in the project, "
                "it is left untouched and not imported again."
            ),
            'existing_update': (
                "Update existing view contents:\n"
                "- If a drafting view with the same name exists, its detailing is cleared and "
                "replaced with the detailing from the source view.\n"
                "- No additional views are created."
            ),
            'existing_duplicate': (
                "Duplicate views:\n"
                "- A new drafting view is created for each imported view, even if a view with "
                "the same name already exists.\n"
                "- The new view gets a unique name (Revit may append a suffix)."
            ),
        }

    def _set_help(self, key):
        self._help_key = key
        texts = self._help_texts()
        msg = texts.get(key, "Choose an option to see more information.")
        self.help_preview.Text = msg


# ------------------------------------------------------------
# UI HELPERS
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


def pick_single_rvt_file():
    path = forms.pick_file(file_ext='rvt', title="Select RVT file")
    if not path:
        sys.exit()
    return path


def list_drafting_view_names(app, path):
    """Open an RVT just to list drafting view names, then close it."""
    names = []

    try:
        mpath = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(path)
        opts = DB.OpenOptions()
        opts.Audit = False
        opts.DetachFromCentralOption = DB.DetachFromCentralOption.DetachAndPreserveWorksets
        opts.AllowOpeningLocalByWrongUser = True

        src_doc = app.OpenDocumentFile(mpath, opts)
    except Exception as err:
        forms.alert("Could not open file:\n{}\n\n{}".format(path, err))
        return names

    try:
        drafts = DB.FilteredElementCollector(src_doc) \
            .OfClass(DB.ViewDrafting) \
            .WhereElementIsNotElementType() \
            .ToElements()

        for v in drafts:
            if not v.IsTemplate:
                try:
                    names.append(v.Name)
                except Exception:
                    pass
    finally:
        try:
            src_doc.Close(False)
        except Exception:
            pass

    return names


def pick_drafting_views_for_single_file(app, path):
    names = list_drafting_view_names(app, path)
    if not names:
        forms.alert("No drafting views found in:\n{}".format(path))
        sys.exit()

    selected = forms.SelectFromList.show(
        sorted(names),
        multiselect=True,
        title="Select Drafting Views",
        button_name="Import"
    )
    if not selected:
        sys.exit()

    # Use a set for fast membership tests
    return set(selected)


# ------------------------------------------------------------
# COPYING ENGINE
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

        # Try to get a name
        name = ""
        try:
            name = el.Name or ""
        except Exception:
            name = ""

        # Skip viewports and extents
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

        # Skip sheet-ish references
        if el.Category and "sheet" in el.Category.Name.lower():
            continue

        result.append(el.Id)

    return result


def find_dest_view(dest_doc, name):
    col = DB.FilteredElementCollector(dest_doc) \
        .OfClass(DB.ViewDrafting) \
        .WhereElementIsNotElementType()

    for v in col:
        if v.IsTemplate:
            continue
        if v.Name == name:
            return v
    return None


def copy_single_drafting_view(src_doc, src_view, dest_doc, existing_policy):
    name = src_view.Name
    existing = find_dest_view(dest_doc, name)

    # Print what we found before making decisions
    if existing:
        print("  Found existing view: {0}".format(name))

        if existing_policy == 'skip':
            print("    → Skipping (per user option)")
            return

        elif existing_policy == 'update':
            print("    → Updating existing view contents")
        elif existing_policy == 'duplicate':
            print("    → Duplicating (Revit will append a suffix)")
    else:
        print("  No existing view found: {0}".format(name))
        print("    → Creating new drafting view")


    # ----------------------------------------------------------
    # CREATE / PREPARE TARGET VIEW
    # ----------------------------------------------------------
    with revit.Transaction("Create/Prepare Drafting View", doc=dest_doc):

        # UPDATE existing
        if existing and existing_policy == 'update':
            target = existing

            # Clear only the detail items
            ids_to_delete = get_detail_elements(dest_doc, target)
            for eid in ids_to_delete:
                try:
                    dest_doc.Delete(eid)
                except Exception:
                    pass

        else:
            # DUPLICATE or CREATE NEW
            vtype = dest_doc.GetDefaultElementTypeId(DB.ElementTypeGroup.ViewTypeDrafting)
            target = DB.ViewDrafting.Create(dest_doc, vtype)

            try:
                target.Name = name   # Revit auto-appends "(2)", "(3)" as needed
            except Exception:
                pass

        # Copy display properties
        try:
            target.Scale = src_view.Scale
        except Exception:
            pass

        try:
            sp = src_view.Parameter[VIEW_TOS_PARAM]
            dp = target.Parameter[VIEW_TOS_PARAM]
            if sp and dp and not dp.IsReadOnly:
                dp.Set(sp.AsString())
        except Exception:
            pass

    # ----------------------------------------------------------
    # COPY DETAIL ELEMENTS
    # ----------------------------------------------------------
    ids = get_detail_elements(src_doc, src_view)
    if not ids:
        print("    → No detail items; skipping: {0}".format(name))
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

    # Final confirmation
    if existing and existing_policy == 'update':
        print("    ✔ Updated existing: {0}".format(target.Name))
    elif existing and existing_policy == 'duplicate':
        print("    ✔ Duplicated as: {0}".format(target.Name))
    elif not existing:
        print("    ✔ Created new drafting view: {0}".format(target.Name))



# ------------------------------------------------------------
# IMPORT RVT FILE(S)
# ------------------------------------------------------------
def import_from_rvt(app, dest_doc, path, existing_policy, view_name_filter=None):
    """Import drafting views from a single RVT.

    view_name_filter:
        - None  => all drafting views
        - set([...]) of names => only those names
    """
    print("\nSOURCE: {0}".format(os.path.basename(path)))

    # Open source doc without activating it
    try:
        mpath = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(path)
        opts = DB.OpenOptions()
        opts.Audit = False
        opts.DetachFromCentralOption = DB.DetachFromCentralOption.DetachAndPreserveWorksets
        opts.AllowOpeningLocalByWrongUser = True

        src_doc = app.OpenDocumentFile(mpath, opts)

    except Exception as err:
        print("  ERROR opening: {0}".format(err))
        return

    try:
        drafts = DB.FilteredElementCollector(src_doc) \
            .OfClass(DB.ViewDrafting) \
            .WhereElementIsNotElementType() \
            .ToElements()

        # Filter non-templates
        drafts = [v for v in drafts if not v.IsTemplate]

        # Optional view filter (single-file mode)
        if view_name_filter is not None:
            filtered = []
            for v in drafts:
                try:
                    if v.Name in view_name_filter:
                        filtered.append(v)
                except Exception:
                    pass
            drafts = filtered

        print("  Drafting views to import: {0}".format(len(drafts)))

        for dv in drafts:
            copy_single_drafting_view(src_doc, dv, dest_doc, existing_policy)

    finally:
        try:
            src_doc.Close(False)
        except Exception:
            pass


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------
def main():
    dest_doc = revit.doc
    app = dest_doc.Application

    # 1) Show options window
    options_win = DraftingImportOptionsWindow()
    options_win.show_dialog()

    mode = options_win.mode              # 'single' or 'multi'
    existing_policy = options_win.existing_policy  # 'skip', 'update', 'duplicate'

    # 2) Decide workflow based on mode
    if mode == 'single':
        # Single RVT file, then pick drafting views from that file
        rvt_path = pick_single_rvt_file()
        selected_view_names = pick_drafting_views_for_single_file(app, rvt_path)

        tg = DB.TransactionGroup(dest_doc, "Import Drafting Views (Single File)")
        tg.Start()
        try:
            import_from_rvt(app, dest_doc, rvt_path, existing_policy, selected_view_names)
            tg.Assimilate()
        except Exception:
            tg.RollBack()
            raise

    else:
        # Multiple files: folder → pick RVTs → import all drafting views in each
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
                import_from_rvt(app, dest_doc, fp, existing_policy, None)
            tg.Assimilate()
        except Exception:
            tg.RollBack()
            raise

    print("\nDone.\n")


main()
