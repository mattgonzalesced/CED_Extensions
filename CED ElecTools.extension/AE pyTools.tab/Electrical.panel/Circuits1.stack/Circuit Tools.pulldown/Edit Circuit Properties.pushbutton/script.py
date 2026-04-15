# -*- coding: utf-8 -*-

import os

from pyrevit import DB, forms, revit

from UIClasses import pathing as ui_pathing

_THIS_DIR = os.path.abspath(os.path.dirname(__file__))
_LIB_ROOT = ui_pathing.ensure_lib_root_on_syspath(_THIS_DIR)
if not _LIB_ROOT:
    forms.alert("Could not locate CEDLib.lib.", title="Edit Circuit Properties", exitscript=True)

from CEDElectrical.Application.dto.operation_request import OperationRequest
from CEDElectrical.Application.services.operation_runner import build_default_runner
from CEDElectrical.Domain import settings_manager
from CEDElectrical.ui.circuit_properties_editor import CircuitPropertiesEditorWindow
from Snippets import _elecutils as eu
from Snippets import revit_helpers
from UIClasses import Resources as UIResources
from UIClasses import load_theme_state_from_config


TITLE = "Edit Circuit Properties"
ALERT_DATA_PARAM = "Circuit Data_CED"


def _idval(item):
    return int(revit_helpers.get_elementid_value(item))


def _electrical_panel_root(start_dir):
    current = os.path.abspath(start_dir)
    while True:
        if os.path.basename(current) == "Electrical.panel":
            return current
        parent = os.path.dirname(current)
        if not parent or parent == current:
            return None
        current = parent


def _editor_xaml_path():
    panel_root = _electrical_panel_root(_THIS_DIR)
    if not panel_root:
        return None
    return os.path.abspath(
        os.path.join(
            panel_root,
            "Circuit Manager.pushbutton",
            "CircuitEditPropertiesWindow.xaml",
        )
    )


def _collect_target_circuits(doc):
    selection = list(revit.get_selection() or [])
    selected = [x for x in selection if isinstance(x, DB.Electrical.ElectricalSystem)]
    if not selected and selection:
        selected = list(eu.get_circuits_from_selection(selection) or [])
    if not selected:
        selected = list(eu.pick_circuits_from_list(doc, select_multiple=True, include_spares_and_spaces=False) or [])
    selected = [x for x in selected if isinstance(x, DB.Electrical.ElectricalSystem)]

    unique = []
    seen = set()
    for circuit in list(selected or []):
        cid = _idval(circuit.Id)
        if cid <= 0 or cid in seen:
            continue
        seen.add(cid)
        unique.append(circuit)
    return unique


def _derive_branch_type(circuit):
    try:
        param = circuit.LookupParameter("CKT_Circuit Type_CEDT")
    except Exception:
        param = None
    if param is not None:
        try:
            value = param.AsString() or param.AsValueString()
        except Exception:
            value = None
        text = str(value or "").strip().upper()
        if text:
            return text

    ctype = getattr(circuit, "CircuitType", None)
    if ctype == DB.Electrical.CircuitType.Space:
        return "SPACE"
    if ctype == DB.Electrical.CircuitType.Spare:
        return "SPARE"
    return "BRANCH"


class _EditorTarget(object):
    def __init__(self, circuit):
        self.circuit = circuit
        panel_name = "-"
        try:
            if circuit.BaseEquipment is not None:
                panel_name = getattr(circuit.BaseEquipment, "Name", panel_name) or panel_name
        except Exception:
            panel_name = "-"
        self.panel = panel_name
        self.circuit_number = getattr(circuit, "CircuitNumber", "") or ""
        self.load_name = getattr(circuit, "LoadName", "") or ""
        self.branch_type = _derive_branch_type(circuit)


def _to_editor_targets(circuits):
    return [_EditorTarget(circuit) for circuit in list(circuits or []) if circuit is not None]


def _theme_state():
    theme_mode, accent_mode = load_theme_state_from_config(
        default_theme="light",
        default_accent="blue",
    )
    return theme_mode, accent_mode


def _open_editor(doc, targets):
    xaml_path = _editor_xaml_path()
    if not xaml_path or not os.path.exists(xaml_path):
        forms.alert("Editor XAML not found.\n\n{}".format(xaml_path or "<missing>"), title=TITLE, exitscript=True)

    settings = settings_manager.load_circuit_settings(doc)
    theme_mode, accent_mode = _theme_state()
    resources_root = (
        UIResources.get_resources_root()
        or ui_pathing.resolve_ui_resources_root(_LIB_ROOT)
        or os.path.abspath(os.path.join(_LIB_ROOT, "UIClasses", "Resources"))
    )

    window = CircuitPropertiesEditorWindow(
        xaml_path=xaml_path,
        targets=targets,
        settings=settings,
        theme_mode=theme_mode,
        accent_mode=accent_mode,
        resources_root=resources_root,
    )
    window.ShowDialog()
    return window


def _run_apply_operation(doc, updates):
    circuit_ids = []
    for row in list(updates or []):
        try:
            cid = int((row or {}).get("circuit_id") or 0)
        except Exception:
            cid = 0
        if cid > 0:
            circuit_ids.append(cid)
    if not circuit_ids:
        forms.alert("No valid circuit updates found.", title=TITLE, exitscript=True)

    request = OperationRequest(
        operation_key="edit_circuit_properties_and_recalculate",
        circuit_ids=circuit_ids,
        source="ribbon",
        options={
            "updates": list(updates or []),
            "show_output": False,
        },
    )
    runner = build_default_runner(alert_parameter_name=ALERT_DATA_PARAM)
    return runner.run(request, doc) or {}


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("Open a model document first.", title=TITLE, exitscript=True)

    circuits = _collect_target_circuits(doc)
    if not circuits:
        forms.alert("No circuits selected.", title=TITLE, exitscript=True)

    targets = _to_editor_targets(circuits)
    if not targets:
        forms.alert("No valid circuit targets selected.", title=TITLE, exitscript=True)

    window = _open_editor(doc, targets)
    if not bool(getattr(window, "apply_requested", False)):
        return

    updates = list((getattr(window, "apply_payload", {}) or {}).get("updates") or [])
    if not updates:
        forms.alert("No staged changes to apply.", title=TITLE, exitscript=True)

    result = _run_apply_operation(doc, updates)
    if result.get("status") == "ok":
        edited = int(result.get("edited_circuits", 0) or 0)
        updated = int(result.get("updated_circuits", 0) or 0)
        forms.alert(
            "Applied staged edits to {} circuit(s).\nRecalculated {} circuit(s).".format(edited, updated),
            title=TITLE,
        )
        return

    reason = result.get("reason", "unknown")
    forms.alert("Edit operation ended: {}".format(reason), title=TITLE)


main()
