# -*- coding: utf-8 -*-
"""
Circuit Fields Audit
--------------------
Filter Electrical Fixtures by Element Linker, Circuit Number, and Panel fields.
Matches are highlighted in orange until cleared.
"""

import json
import os
from pyrevit import forms, revit, script
from Autodesk.Revit.DB import (
    BuiltInCategory,
    Color,
    ElementId,
    FilteredElementCollector,
    FillPatternElement,
    OverrideGraphicSettings,
    StorageType,
    Transaction,
)
from Autodesk.Revit.DB.ExtensibleStorage import Schema, SchemaBuilder, Entity
from System import Guid, String


TITLE = "Circuit Fields Audit"
SCHEMA_GUID = Guid("e7d54f95-1c5d-4df7-95fd-8a1b8be727a8")
SCHEMA_NAME = "CED_CircuitFieldsAudit"
DATA_FIELD = "DataJson"

LINKER_PARAM_NAMES = ("Element_Linker Parameter", "Element_Linker")
CKT_CIRCUIT_PARAM = "CKT_Circuit Number_CEDT"
CKT_PANEL_PARAM = "CKT_Panel_CEDT"

ORANGE_RGB = (255, 128, 0)

output = script.get_output()
output.close_others()


def _get_schema():
    schema = Schema.Lookup(SCHEMA_GUID)
    if schema:
        return schema
    builder = SchemaBuilder(SCHEMA_GUID)
    builder.SetSchemaName(SCHEMA_NAME)
    builder.AddSimpleField(DATA_FIELD, String)
    return builder.Finish()


def _read_state(doc):
    project_info = getattr(doc, "ProjectInformation", None)
    if project_info is None:
        return {}
    schema = _get_schema()
    entity = project_info.GetEntity(schema)
    if not entity or not entity.IsValid():
        return {}
    field = schema.GetField(DATA_FIELD)
    raw = None
    try:
        raw = entity.Get[String](field)
    except Exception:
        try:
            raw = entity.Get[str](field)
        except Exception:
            raw = None
    if not raw:
        return {}
    try:
        data = json.loads(raw)
    except Exception:
        return {}
    return data if isinstance(data, dict) else {}


def _write_state(doc, data):
    project_info = getattr(doc, "ProjectInformation", None)
    if project_info is None:
        return
    schema = _get_schema()
    field = schema.GetField(DATA_FIELD)
    payload = json.dumps(data or {})
    entity = Entity(schema)
    try:
        entity.Set[String](field, payload)
    except Exception:
        entity.Set[str](field, payload)
    project_info.SetEntity(entity)


def _get_solid_fill_pattern_id(doc):
    for fpe in FilteredElementCollector(doc).OfClass(FillPatternElement):
        try:
            fp = fpe.GetFillPattern()
            if fp and fp.IsSolidFill:
                return fpe.Id
        except Exception:
            pass
    return ElementId.InvalidElementId


def _build_orange_overrides(doc):
    ogs = OverrideGraphicSettings()
    r, g, b = ORANGE_RGB
    col = Color(bytearray([r])[0], bytearray([g])[0], bytearray([b])[0])
    try:
        ogs.SetProjectionLineColor(col)
    except Exception:
        pass
    try:
        ogs.SetCutLineColor(col)
    except Exception:
        pass
    solid_id = _get_solid_fill_pattern_id(doc)
    if solid_id and solid_id != ElementId.InvalidElementId:
        try:
            ogs.SetSurfaceForegroundPatternId(solid_id)
            ogs.SetSurfaceForegroundPatternColor(col)
        except Exception:
            pass
        try:
            ogs.SetCutForegroundPatternId(solid_id)
            ogs.SetCutForegroundPatternColor(col)
        except Exception:
            pass
    return ogs


def _lookup_param(elem, name):
    try:
        return elem.LookupParameter(name)
    except Exception:
        return None


def _param_to_text(param):
    if param is None:
        return None
    try:
        if param.StorageType == StorageType.String:
            value = param.AsString()
            if value is not None:
                return value
    except Exception:
        pass
    try:
        value = param.AsString()
        if value is not None:
            return value
    except Exception:
        pass
    try:
        value = param.AsValueString()
        if value is not None:
            return value
    except Exception:
        pass
    try:
        if param.StorageType == StorageType.Integer:
            return str(param.AsInteger())
    except Exception:
        pass
    try:
        if param.StorageType == StorageType.Double:
            return str(param.AsDouble())
    except Exception:
        pass
    try:
        if param.StorageType == StorageType.ElementId:
            elem_id = param.AsElementId()
            if elem_id and elem_id.IntegerValue != -1:
                return str(elem_id.IntegerValue)
    except Exception:
        pass
    return None


def _get_param_value(elem, names):
    if not isinstance(names, (list, tuple)):
        names = [names]
    for name in names:
        param = _lookup_param(elem, name)
        if param is None:
            try:
                symbol = getattr(elem, "Symbol", None)
                if symbol is not None:
                    param = symbol.LookupParameter(name)
            except Exception:
                param = None
        if param is None:
            continue
        value = _param_to_text(param)
        if value is not None:
            return value
    return None


def _passes(value, search_text, allow_blank):
    text = (search_text or "").strip().lower()
    value_norm = (value or "").strip()
    if not text and not allow_blank:
        return True
    if allow_blank and not value_norm:
        return True
    if text and value_norm and text in value_norm.lower():
        return True
    return False


def _collect_matches(doc, criteria):
    collector = FilteredElementCollector(doc)
    collector = collector.OfCategory(BuiltInCategory.OST_ElectricalFixtures)
    collector = collector.WhereElementIsNotElementType()
    matches = []
    for elem in collector:
        linker_val = _get_param_value(elem, LINKER_PARAM_NAMES)
        circuit_val = _get_param_value(elem, CKT_CIRCUIT_PARAM)
        panel_val = _get_param_value(elem, CKT_PANEL_PARAM)
        if not _passes(linker_val, criteria.get("linker_text"), criteria.get("linker_blank")):
            continue
        if not _passes(circuit_val, criteria.get("circuit_text"), criteria.get("circuit_blank")):
            continue
        if not _passes(panel_val, criteria.get("panel_text"), criteria.get("panel_blank")):
            continue
        matches.append(elem)
    return matches


def _apply_overrides(view, element_ids, ogs):
    if view is None:
        return
    for elem_id in element_ids:
        try:
            view.SetElementOverrides(ElementId(int(elem_id)), ogs)
        except Exception:
            continue


def _clear_overrides(view, element_ids):
    if view is None:
        return
    ogs = OverrideGraphicSettings()
    for elem_id in element_ids:
        try:
            view.SetElementOverrides(ElementId(int(elem_id)), ogs)
        except Exception:
            continue


def _resolve_view(doc, view_id):
    if view_id is None:
        return None
    try:
        view_elem = doc.GetElement(ElementId(int(view_id)))
        return view_elem
    except Exception:
        return None


def _print_results(element_ids, criteria):
    output.print_md("### Circuit Fields Audit")
    output.print_md("Filtered Electrical Fixtures: `{}`".format(len(element_ids)))
    output.print_md("")
    output.print_md("**Filters**")
    output.print_md("- Element Linker: text='{}' blank={}".format(
        criteria.get("linker_text") or "",
        bool(criteria.get("linker_blank")),
    ))
    output.print_md("- CKT_Circuit Number_CEDT: text='{}' blank={}".format(
        criteria.get("circuit_text") or "",
        bool(criteria.get("circuit_blank")),
    ))
    output.print_md("- CKT_Panel_CEDT: text='{}' blank={}".format(
        criteria.get("panel_text") or "",
        bool(criteria.get("panel_blank")),
    ))
    output.print_md("")
    if not element_ids:
        output.print_md("No matching fixtures found.")
        return
    output.print_md("**ElementIds**")
    for elem_id in element_ids:
        output.print_md("- `{}`".format(elem_id))


class CircuitFieldsAuditWindow(forms.WPFWindow):
    def __init__(self, xaml_path, doc):
        forms.WPFWindow.__init__(self, xaml_path)
        self._doc = doc

        self._element_linker_text = self.FindName("ElementLinkerText")
        self._element_linker_blank = self.FindName("ElementLinkerBlank")
        self._circuit_text = self.FindName("CircuitNumberText")
        self._circuit_blank = self.FindName("CircuitNumberBlank")
        self._panel_text = self.FindName("PanelText")
        self._panel_blank = self.FindName("PanelBlank")

        self._apply_btn = self.FindName("ApplyButton")
        self._clear_btn = self.FindName("ClearButton")
        self._close_btn = self.FindName("CloseButton")
        self._status = self.FindName("StatusText")

        if self._apply_btn is not None:
            self._apply_btn.Click += self._on_apply
        if self._clear_btn is not None:
            self._clear_btn.Click += self._on_clear
        if self._close_btn is not None:
            self._close_btn.Click += self._on_close

        self._set_status(self._initial_status())

    def _initial_status(self):
        data = _read_state(self._doc)
        stored = data.get("element_ids") or []
        if stored:
            return "Stored highlight: {} element(s).".format(len(stored))
        return "Ready."

    def _set_status(self, text):
        if self._status is not None:
            self._status.Text = text

    def _collect_criteria(self):
        return {
            "linker_text": self._element_linker_text.Text if self._element_linker_text is not None else "",
            "linker_blank": bool(self._element_linker_blank.IsChecked) if self._element_linker_blank is not None else False,
            "circuit_text": self._circuit_text.Text if self._circuit_text is not None else "",
            "circuit_blank": bool(self._circuit_blank.IsChecked) if self._circuit_blank is not None else False,
            "panel_text": self._panel_text.Text if self._panel_text is not None else "",
            "panel_blank": bool(self._panel_blank.IsChecked) if self._panel_blank is not None else False,
        }

    def _on_apply(self, sender, args):
        active_view = revit.active_view
        if active_view is None:
            forms.alert("No active view found.", title=TITLE)
            return

        criteria = self._collect_criteria()
        matches = _collect_matches(self._doc, criteria)
        element_ids = [elem.Id.IntegerValue for elem in matches]
        element_ids = sorted(set(element_ids))

        prev_state = _read_state(self._doc)
        prev_ids = prev_state.get("element_ids") or []
        prev_view_id = prev_state.get("view_id")
        prev_view = _resolve_view(self._doc, prev_view_id) or active_view

        ogs = _build_orange_overrides(self._doc)

        tx = Transaction(self._doc, "Circuit Fields Audit")
        tx.Start()
        try:
            if prev_ids:
                _clear_overrides(prev_view, prev_ids)
            _apply_overrides(active_view, element_ids, ogs)
            _write_state(self._doc, {
                "element_ids": element_ids,
                "view_id": active_view.Id.IntegerValue,
            })
            tx.Commit()
        except Exception as ex:
            tx.RollBack()
            forms.alert("Failed to apply filters:\n\n{}".format(ex), title=TITLE)
            return

        _print_results(element_ids, criteria)
        self._set_status("Highlighted {} element(s) in view '{}'.".format(
            len(element_ids),
            getattr(active_view, "Name", "<view>")
        ))

    def _on_clear(self, sender, args):
        active_view = revit.active_view
        if active_view is None:
            forms.alert("No active view found.", title=TITLE)
            return
        state = _read_state(self._doc)
        element_ids = state.get("element_ids") or []
        view_id = state.get("view_id")
        target_view = _resolve_view(self._doc, view_id) or active_view
        if not element_ids:
            self._set_status("No stored highlight to clear.")
            return

        tx = Transaction(self._doc, "Circuit Fields Audit - Clear")
        tx.Start()
        try:
            _clear_overrides(target_view, element_ids)
            _write_state(self._doc, {
                "element_ids": [],
                "view_id": None,
            })
            tx.Commit()
        except Exception as ex:
            tx.RollBack()
            forms.alert("Failed to clear filters:\n\n{}".format(ex), title=TITLE)
            return

        self._set_status("Filters cleared.")

    def _on_close(self, sender, args):
        self.Close()


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document detected.", title=TITLE)
        return

    xaml_path = os.path.join(os.path.dirname(__file__), "CircuitFieldsAuditWindow.xaml")
    if not os.path.exists(xaml_path):
        forms.alert("Missing UI file: {}".format(xaml_path), title=TITLE)
        return

    window = CircuitFieldsAuditWindow(xaml_path, doc)
    window.ShowDialog()


if __name__ == "__main__":
    main()
