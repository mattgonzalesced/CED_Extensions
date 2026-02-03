# -*- coding: utf-8 -*-
__title__ = "System Tagger"
__doc__ = "Place system ID text notes on refrigerated cases from an Excel list."

import os

from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import revit, DB, forms, script
from pyrevit.interop import xl as pyxl


logger = script.get_logger()
doc = revit.doc
uidoc = revit.uidoc

ORANGE_RGB = (255, 128, 0)
CONFIG_NAME = "system_tagger_config"


def _rgb(r, g, b):
    return DB.Color(bytearray([r])[0], bytearray([g])[0], bytearray([b])[0])


def _get_config():
    return script.get_config(CONFIG_NAME)


def _load_resume_state():
    cfg = _get_config()
    path = getattr(cfg, "resume_excel_path", None)
    idx = getattr(cfg, "resume_index", None)
    try:
        idx = int(idx)
    except Exception:
        idx = None
    return path, idx


def _save_resume_state(path, idx):
    cfg = _get_config()
    cfg.resume_excel_path = path
    cfg.resume_index = int(idx)
    script.save_config()


def _clear_resume_state():
    cfg = _get_config()
    cfg.resume_excel_path = None
    cfg.resume_index = None
    script.save_config()


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


def _pick_text_type():
    types = list(DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType))
    if not types:
        return None
    for t in types:
        try:
            name = t.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or ""
            font = t.get_Parameter(DB.BuiltInParameter.TEXT_FONT).AsString() or ""
            if ("3/32" in name) and ("Arial" in font):
                return t
        except Exception:
            continue
    return types[0]


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


def _read_system_ids_from_excel(path):
    data = pyxl.load(path, sheets=None, columns=["System ID"])
    if not data:
        return []
    sheet_name = list(data.keys())[0]
    rows = data[sheet_name]["rows"]
    system_ids = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        raw = row.get("System ID")
        if raw is None:
            continue
        text = str(raw).strip()
        if not text:
            continue
        if text.lower() == "system id":
            continue
        system_ids.append(text)
    return system_ids


def _toggle_pick_cases(system_id, view, ogs):
    selected_ids = []
    selected_set = set()

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
                    with revit.Transaction("System Tagger - Unhighlight"):
                        _clear_override(view, elem_id)
                else:
                    selected_set.add(elem_int)
                    selected_ids.append(elem_int)
                    with revit.Transaction("System Tagger - Highlight"):
                        _apply_override(view, elem_id, ogs)

        choice = forms.CommandSwitchWindow.show(
            ["Done", "Pick More", "Pause", "Cancel"],
            message="System ID '{}': {} case(s) selected.".format(system_id, len(selected_ids)),
        )
        if choice == "Pick More":
            continue
        if choice == "Pause":
            return "pause", selected_ids
        if choice == "Cancel" or choice is None:
            return "cancel", selected_ids
        return "done", selected_ids


def _place_text_notes(system_id, element_ids, view, text_type):
    if not element_ids:
        return 0
    if len(element_ids) == 1:
        labels = [system_id]
    else:
        labels = [system_id + _index_to_letters(i) for i in range(len(element_ids))]

    count = 0
    with revit.Transaction("System Tagger - Place Text Notes"):
        for elem_int, label in zip(element_ids, labels):
            elem = doc.GetElement(DB.ElementId(elem_int))
            if not elem:
                continue
            pt = _get_element_center(elem, view)
            if not pt:
                logger.warning("No bounding box for element {}".format(elem_int))
                continue
            opts = DB.TextNoteOptions()
            opts.TypeId = text_type.Id
            try:
                opts.HorizontalAlignment = DB.HorizontalTextAlignment.Center
            except Exception:
                pass
            try:
                opts.VerticalAlignment = DB.VerticalTextAlignment.Middle
            except Exception:
                pass
            try:
                DB.TextNote.Create(doc, view.Id, pt, label, opts)
                count += 1
            except Exception as ex:
                logger.warning("Failed to place text note for {}: {}".format(elem_int, ex))
    return count


def _clear_highlights(view, element_ids):
    if not element_ids:
        return
    with revit.Transaction("System Tagger - Clear Highlights"):
        clear_ogs = DB.OverrideGraphicSettings()
        for elem_int in element_ids:
            view.SetElementOverrides(DB.ElementId(elem_int), clear_ogs)


def main():
    active_view = revit.active_view
    if active_view.IsTemplate:
        forms.alert("Active view is a template. Open a working view first.", exitscript=True)

    resume_path, resume_index = _load_resume_state()
    system_ids = None
    start_index = 0
    path = None

    if resume_path and os.path.exists(resume_path) and resume_index is not None:
        try:
            resume_ids = _read_system_ids_from_excel(resume_path)
        except Exception:
            resume_ids = None
        if resume_ids and 0 <= resume_index < len(resume_ids):
            next_id = resume_ids[resume_index]
            choice = forms.CommandSwitchWindow.show(
                ["Resume", "Start Over", "Cancel"],
                message="Resume System Tagger at '{}'? ({} of {})".format(
                    next_id, resume_index + 1, len(resume_ids)
                ),
            )
            if choice == "Resume":
                path = resume_path
                system_ids = resume_ids
                start_index = resume_index
            elif choice == "Start Over":
                _clear_resume_state()
            else:
                script.exit()
        else:
            _clear_resume_state()

    if system_ids is None:
        path = forms.pick_file(file_ext="xlsx", title="Select System ID Excel File")
        if not path:
            script.exit()
        try:
            system_ids = _read_system_ids_from_excel(path)
        except Exception as ex:
            forms.alert("Failed to read Excel file: {}".format(ex), exitscript=True)
        if not system_ids:
            forms.alert("No System IDs found in the first column. Expected header: 'System ID'.", exitscript=True)

    text_type = _pick_text_type()
    if not text_type:
        forms.alert("No TextNoteType available in this project.", exitscript=True)

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
        _save_resume_state(path, idx)
        status, picked_ids = _toggle_pick_cases(system_id, active_view, highlight_ogs)
        if status == "cancel":
            _clear_highlights(active_view, picked_ids)
            return
        try:
            _place_text_notes(system_id, picked_ids, active_view, text_type)
            _clear_highlights(active_view, picked_ids)
            _save_resume_state(path, idx + 1)
        except Exception:
            _clear_highlights(active_view, picked_ids)
            raise
        if status == "pause":
            forms.alert("System Tagger paused. Run again to resume.")
            return

    _clear_resume_state()


if __name__ == "__main__":
    main()
