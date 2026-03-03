# -*- coding: utf-8 -*-
__title__ = "System Tagger"
__doc__ = "Place system ID tags on refrigerated cases from a pasted SYS NO. list."

import re
import time

from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import revit, DB, forms, script
from Autodesk.Revit.DB.ExtensibleStorage import Entity, Schema, SchemaBuilder
from Autodesk.Revit.DB.ExtensibleStorage import AccessLevel
from System import Guid, String, Int32


logger = script.get_logger()
doc = revit.doc
uidoc = revit.uidoc

ORANGE_RGB = (255, 128, 0)
RESUME_SCHEMA_GUID = Guid("5f1a3c2e-9e4c-4e42-9f26-9c8f8f5c6f14")
RESUME_SCHEMA_NAME = "CED_SystemTagger_Resume"


def _rgb(r, g, b):
    return DB.Color(bytearray([r])[0], bytearray([g])[0], bytearray([b])[0])


def _load_resume_state():
    if doc is None:
        return None, None
    schema = _get_resume_schema()
    proj_info = doc.ProjectInformation
    if proj_info is None:
        return None, None
    try:
        entity = proj_info.GetEntity(schema)
    except Exception:
        return None, None
    if not entity or not entity.IsValid():
        return None, None
    try:
        path = entity.Get[String](schema.GetField("ExcelPath"))
    except Exception:
        path = None
    try:
        idx = entity.Get[Int32](schema.GetField("Index"))
    except Exception:
        idx = None
    if idx is not None and idx < 0:
        idx = None
    return path or None, idx


def _save_resume_state(path, idx):
    if doc is None:
        return
    schema = _get_resume_schema()
    proj_info = doc.ProjectInformation
    if proj_info is None:
        return
    entity = Entity(schema)
    try:
        entity.Set[String](schema.GetField("ExcelPath"), path or "")
    except Exception:
        pass
    try:
        entity.Set[Int32](schema.GetField("Index"), int(idx))
    except Exception:
        pass
    with revit.Transaction("System Tagger - Save Resume State"):
        proj_info.SetEntity(entity)


def _clear_resume_state():
    _save_resume_state("", -1)


def _get_resume_schema():
    schema = Schema.Lookup(RESUME_SCHEMA_GUID)
    if schema is not None:
        return schema
    sb = SchemaBuilder(RESUME_SCHEMA_GUID)
    sb.SetSchemaName(RESUME_SCHEMA_NAME)
    sb.SetReadAccessLevel(AccessLevel.Public)
    sb.SetWriteAccessLevel(AccessLevel.Public)
    sb.AddSimpleField("ExcelPath", String)
    sb.AddSimpleField("Index", Int32)
    return sb.Finish()


def _get_solid_fill_pattern_id():
    for fpe in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement):
        try:
            fp = fpe.GetFillPattern()
            if fp.Target == DB.FillPatternTarget.Drafting and fp.IsSolidFill:
                return fpe.Id
        except Exception:
            continue
    return None


def _build_highlight_ogs():
    ogs = DB.OverrideGraphicSettings()
    col = _rgb(*ORANGE_RGB)
    try:
        ogs.SetProjectionLineColor(col)
    except Exception:
        pass
    try:
        ogs.SetCutLineColor(col)
    except Exception:
        pass
    pattern_id = _get_solid_fill_pattern_id()
    if pattern_id:
        try:
            ogs.SetSurfaceForegroundPatternId(pattern_id)
            ogs.SetSurfaceForegroundPatternColor(col)
        except Exception:
            pass
    return ogs


def _clear_override(view, elem_id):
    view.SetElementOverrides(elem_id, DB.OverrideGraphicSettings())


def _apply_override(view, elem_id, ogs):
    view.SetElementOverrides(elem_id, ogs)


def _apply_highlights(view, element_ids, ogs):
    if not element_ids:
        return
    for elem_int in element_ids:
        _apply_override(view, DB.ElementId(elem_int), ogs)


def _clear_highlights_in_tx(view, element_ids):
    if not element_ids:
        return
    clear_ogs = DB.OverrideGraphicSettings()
    for elem_int in element_ids:
        view.SetElementOverrides(DB.ElementId(elem_int), clear_ogs)


def _update_preview_highlights(view, ogs, new_ids, prev_ids):
    try:
        with revit.Transaction("System Tagger - Preview Highlights"):
            _clear_highlights_in_tx(view, prev_ids)
            _apply_highlights(view, new_ids, ogs)
        return list(new_ids or [])
    except Exception as ex:
        logger.warning("Preview highlight update failed: {}".format(ex))
        return list(prev_ids or [])




class SystemIdPasteWindow(forms.WPFWindow):
    def __init__(self, xaml_path):
        forms.WPFWindow.__init__(self, xaml_path)

    def OkButton_Click(self, sender, args):
        self.DialogResult = True
        self.Close()

    def CancelButton_Click(self, sender, args):
        self.DialogResult = False
        self.Close()

    def get_text(self):
        try:
            return (self.InputBox.Text or u"").strip()
        except Exception:
            return u""



def _prompt_system_ids():
    xaml = script.get_bundle_file("SystemIdPasteWindow.xaml")
    if xaml:
        window = SystemIdPasteWindow(xaml)
        if window.show_dialog():
            return window.get_text()
        return None
    return forms.ask_for_string(
        prompt="Paste SYS NO. values from Excel (full block is OK).",
        default="",
        title="System Tagger"
    )


def _get_element_center(elem, view):
    if not elem:
        return None
    bbox = None
    try:
        bbox = elem.get_BoundingBox(view)
    except Exception:
        bbox = None
    if not bbox:
        try:
            bbox = elem.get_BoundingBox(None)
        except Exception:
            bbox = None
    if not bbox:
        return None
    return (bbox.Min + bbox.Max) * 0.5


def _index_to_letters(index):
    letters = ""
    num = index
    while True:
        num, rem = divmod(num, 26)
        letters = chr(ord("A") + rem) + letters
        if num == 0:
            break
        num -= 1
    return letters


def _collect_tag_types():
    tag_types = []
    for cat_name in ("OST_MultiCategoryTags", "OST_MechanicalEquipmentTags", "OST_SpecialityEquipmentTags"):
        try:
            cat = getattr(DB.BuiltInCategory, cat_name)
        except Exception:
            cat = None
        if cat is None:
            continue
        tag_types.extend(
            DB.FilteredElementCollector(doc)
            .OfClass(DB.FamilySymbol)
            .OfCategory(cat)
            .ToElements()
        )
    return tag_types


def _tag_type_label(tag_type):
    try:
        fam_name = tag_type.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString()
    except Exception:
        fam_name = None
    try:
        type_name = tag_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    except Exception:
        type_name = None
    try:
        cat_name = tag_type.Category.Name if tag_type.Category else "Tag"
    except Exception:
        cat_name = "Tag"
    return "[{}] {} : {}".format(cat_name, fam_name or "?", type_name or "?")


def _pick_tag_type():
    tag_types = _collect_tag_types()
    if not tag_types:
        return None
    for tag_type in tag_types:
        try:
            fam_name = tag_type.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString() or ""
        except Exception:
            fam_name = ""
        try:
            type_name = tag_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or ""
        except Exception:
            type_name = ""
        if fam_name == "M_Mechanical Eqmt Tag" and type_name == "Identity":
            return tag_type
    options = [_tag_type_label(t) for t in tag_types]
    picked = forms.SelectFromList.show(
        options,
        multiselect=False,
        title="Select Tag Type",
    )
    if not picked:
        return None
    try:
        idx = options.index(picked)
    except Exception:
        return None
    return tag_types[idx]


def _parse_system_ids(raw_text):
    system_ids = []
    if not raw_text:
        return system_ids
    raw = raw_text.replace("\r", "\n").strip()
    if not raw:
        return system_ids
    tokens = [t.strip() for t in re.split(r"[\t,\n;]+", raw) if t.strip()]
    for token in tokens:
        lowered = token.lower()
        if lowered in ("system id", "sys no.", "sys no", "sys no#", "sys no #"):
            continue
        system_ids.append(token)
    return system_ids


def _toggle_pick_cases(system_id, view, ogs):
    selected_ids = []
    selected_set = set()
    preview_ids = []
    last_preview_time = 0.0
    preview_group = DB.TransactionGroup(doc, "System Tagger - Preview Highlights")
    preview_group.Start()

    try:
        while True:
            with forms.WarningBar(
                title="System ID {}: pick refrigerated cases (ESC to finish)".format(system_id)
            ):
                while True:
                    try:
                        ref = uidoc.Selection.PickObject(
                            ObjectType.Element,
                            "Pick case (click again to unselect). ESC when done."
                        )
                    except Exception:
                        break

                    elem_int = ref.ElementId.IntegerValue
                    elem_id = DB.ElementId(elem_int)

                    if elem_int in selected_set:
                        selected_set.remove(elem_int)
                        selected_ids = [i for i in selected_ids if i != elem_int]
                    else:
                        selected_set.add(elem_int)
                        selected_ids.append(elem_int)

                    now = time.time()
                    if now - last_preview_time >= 1.0:
                        preview_ids = _update_preview_highlights(view, ogs, selected_ids, preview_ids)
                        last_preview_time = now

            choice = forms.CommandSwitchWindow.show(
                ["Done", "Pick More", "Pause"],
                message="System ID '{}': {} case(s) selected.".format(system_id, len(selected_ids)),
            )
            if choice == "Pick More":
                continue
            if choice == "Pause":
                preview_ids = _update_preview_highlights(view, ogs, selected_ids, preview_ids)
                _update_preview_highlights(view, ogs, [], preview_ids)
                try:
                    preview_group.Assimilate()
                except Exception:
                    pass
                return "pause", selected_ids
            if choice == "Cancel" or choice is None:
                preview_ids = _update_preview_highlights(view, ogs, selected_ids, preview_ids)
                _update_preview_highlights(view, ogs, [], preview_ids)
                try:
                    preview_group.Assimilate()
                except Exception:
                    pass
                return "cancel", selected_ids
            preview_ids = _update_preview_highlights(view, ogs, selected_ids, preview_ids)
            try:
                preview_group.Assimilate()
            except Exception:
                pass
            return "done", selected_ids
    finally:
        try:
            if preview_group.HasStarted():
                preview_group.RollBack()
        except Exception:
            pass
        try:
            preview_group.Dispose()
        except Exception:
            pass


def _build_labels(system_id, element_ids):
    if not element_ids:
        return []
    if len(element_ids) == 1:
        labels = [system_id]
    else:
        labels = [system_id + _index_to_letters(i) for i in range(len(element_ids))]
    return labels

def _apply_identity_mark(element_ids, labels):
    if not element_ids or not labels:
        return
    for elem_int, label in zip(element_ids, labels):
        elem = doc.GetElement(DB.ElementId(elem_int))
        if not elem:
            continue
        try:
            param = elem.LookupParameter("Identity Mark")
            if not param:
                param = elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
            if param and not param.IsReadOnly:
                if param.StorageType == DB.StorageType.String:
                    param.Set(label)
                else:
                    param.SetValueString(label)
        except Exception as ex:
            logger.warning(
                "Failed to set Identity Mark for {}: {}".format(elem_int, ex)
            )


def _place_tags(element_ids, labels, view, tag_type):
    if not element_ids or not labels or tag_type is None:
        return 0
    if not tag_type.IsActive:
        tag_type.Activate()
        doc.Regenerate()
    count = 0
    for elem_int, label in zip(element_ids, labels):
        elem = doc.GetElement(DB.ElementId(elem_int))
        if not elem:
            continue
        pt = _get_element_center(elem, view)
        if not pt:
            logger.warning("No bounding box for element {}".format(elem_int))
            continue
        try:
            reference = DB.Reference(elem)
            DB.IndependentTag.Create(
                doc,
                tag_type.Id,
                view.Id,
                reference,
                False,
                DB.TagOrientation.Horizontal,
                pt,
            )
            count += 1
        except Exception as ex:
            logger.warning("Failed to place tag for {}: {}".format(elem_int, ex))
    return count


def main():
    active_view = revit.active_view
    if active_view.IsTemplate:
        forms.alert("Active view is a template. Open a working view first.", exitscript=True)

    tag_type = _pick_tag_type()
    if not tag_type:
        forms.alert("No tag type available in this project.", exitscript=True)

    resume_path, resume_index = _load_resume_state()
    system_ids = None
    start_index = 0
    raw_text = None

    if resume_path and resume_index is not None:
        resume_ids = _parse_system_ids(resume_path)
        if resume_ids and 0 <= resume_index < len(resume_ids):
            next_id = resume_ids[resume_index]
            choice = forms.CommandSwitchWindow.show(
                ["Resume", "Start Over", "Cancel"],
                message="Resume System Tagger at '{}'? ({} of {})".format(
                    next_id, resume_index + 1, len(resume_ids)
                ),
            )
            if choice == "Resume":
                raw_text = resume_path
                system_ids = resume_ids
                start_index = resume_index
            elif choice == "Start Over":
                _clear_resume_state()
            else:
                script.exit()
        else:
            _clear_resume_state()

    if system_ids is None:
        raw_text = _prompt_system_ids()
        if not raw_text:
            script.exit()
        system_ids = _parse_system_ids(raw_text)
        if not system_ids:
            forms.alert("No System IDs found in the pasted text.", exitscript=True)

    highlight_ogs = _build_highlight_ogs()

    forms.alert(
        "Select refrigerated cases for each System ID.\n"
        "- Click a case again to unselect.\n"
        "- Press ESC when done selecting, then click Done.\n"
        "- Click Pause to save progress and exit.\n"
        "The script will advance to the next System ID."
    )

    for idx in range(start_index, len(system_ids)):
        system_id = system_ids[idx]
        _save_resume_state(raw_text, idx)
        status, picked_ids = _toggle_pick_cases(system_id, active_view, highlight_ogs)
        if status == "cancel":
            return
        try:
            labels = _build_labels(system_id, picked_ids)
            with revit.Transaction("System Tagger - Place Labels {}".format(system_id)):
                _apply_highlights(active_view, picked_ids, highlight_ogs)
                _apply_identity_mark(picked_ids, labels)
                _place_tags(picked_ids, labels, active_view, tag_type)
                _clear_highlights_in_tx(active_view, picked_ids)
            _save_resume_state(raw_text, idx + 1)
        except Exception:
            raise
        if status == "pause":
            forms.alert("System Tagger paused. Run again to resume.")
            return

    _clear_resume_state()


if __name__ == "__main__":
    main()
