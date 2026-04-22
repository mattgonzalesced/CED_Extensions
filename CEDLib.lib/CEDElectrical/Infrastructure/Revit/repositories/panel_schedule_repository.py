# -*- coding: utf-8 -*-
"""Reusable panel schedule helpers for circuit-centric tools."""

import Autodesk.Revit.DB.Electrical as DBE
import clr
from System.Collections.Generic import IList
from System.Collections.Generic import List
from pyrevit import DB, forms

from CEDElectrical.part_types import (
    PART_TYPE_MAP,
    PART_TYPE_OTHER_PANEL,
    PART_TYPE_PANELBOARD,
    PART_TYPE_SWITCHBOARD,
)
from Snippets import revit_helpers

SORT_MODE_SWITCHBOARD = "switchboard"
SORT_MODE_PANELBOARD_ACROSS = "panelboard_two_columns_across"
SORT_MODE_PANELBOARD_DOWN = "panelboard_two_columns_down"
SORT_MODE_PANELBOARD_ONE_COLUMN = "panelboard_one_column"
_PANEL_SORT_MODES = (
    SORT_MODE_PANELBOARD_ACROSS,
    SORT_MODE_PANELBOARD_DOWN,
    SORT_MODE_PANELBOARD_ONE_COLUMN,
)
_DESIGN_OPTION_MAIN_MODEL_FILTER = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)
PSTYPE_UNKNOWN = DBE.PanelScheduleType.Unknown
PSTYPE_BRANCH = DBE.PanelScheduleType.Branch
PSTYPE_SWITCHBOARD = DBE.PanelScheduleType.Switchboard
PSTYPE_DATA = DBE.PanelScheduleType.Data
_PART_TYPE_TO_SCHEDULE_TYPE = {
    PART_TYPE_PANELBOARD: PSTYPE_BRANCH,
    PART_TYPE_SWITCHBOARD: PSTYPE_SWITCHBOARD,
    PART_TYPE_OTHER_PANEL: PSTYPE_DATA,
}


def _distribution_equipment_repo():
    """Lazy import to avoid CEDElectrical.Model import cycles at module load."""
    from . import distribution_equipment_repository

    return distribution_equipment_repository


class PanelLayoutInfo(dict):
    """Lightweight layout descriptor for panel schedule behavior."""

    def __init__(
        self,
        schedule_type,
        panel_configuration,
        sort_mode,
        board_type,
        max_slot=0,
        show_slots_from_device=False,
    ):
        dict.__init__(self)
        self["schedule_type"] = schedule_type
        self["schedule_type_name"] = _to_text(schedule_type, "Unknown")
        self["panel_configuration"] = panel_configuration
        self["panel_configuration_name"] = _to_text(panel_configuration, "Unknown")
        self["sort_mode"] = _to_text(sort_mode, SORT_MODE_PANELBOARD_ACROSS)
        self["board_type"] = _to_text(board_type, "Panelboard")
        self["max_slot"] = int(max_slot or 0)
        self["show_slots_from_device"] = bool(show_slots_from_device)


class PanelEquipmentOption(dict):
    """Lightweight panel context object used by planner UIs."""

    def __init__(self, data):
        dict.__init__(self)
        for key, value in dict(data or {}).items():
            self[key] = value


class PanelScheduleTemplateOption(dict):
    """Lightweight template option used for creating schedule instances."""

    def __init__(self, data):
        dict.__init__(self)
        for key, value in dict(data or {}).items():
            self[key] = value


def _idval(item):
    """Return numeric value for ElementId-like objects."""
    return revit_helpers.get_elementid_value(item)


def _to_text(value, fallback=""):
    """Return a safe text representation."""
    if value is None:
        return fallback
    try:
        return str(value)
    except Exception:
        return fallback


def _is_data_panel_schedule_type(schedule_type):
    """Return True when schedule type corresponds to data panel templates."""
    if schedule_type is None:
        return False
    try:
        if PSTYPE_DATA is not None and schedule_type == PSTYPE_DATA:
            return True
    except Exception:
        pass
    return _to_text(schedule_type, "").strip().lower() == "data"


def _is_switchboard_schedule_type(schedule_type):
    """Return True when schedule type corresponds to switchboard schedules."""
    if schedule_type is None:
        return False
    try:
        if PSTYPE_SWITCHBOARD is not None and schedule_type == PSTYPE_SWITCHBOARD:
            return True
    except Exception:
        pass
    return _to_text(schedule_type, "").strip().lower() == "switchboard"


def _is_unknown_schedule_type(schedule_type):
    """Return True when schedule type is unknown or unavailable."""
    if schedule_type is None:
        return True
    try:
        if PSTYPE_UNKNOWN is not None and schedule_type == PSTYPE_UNKNOWN:
            return True
    except Exception:
        pass
    return _to_text(schedule_type, "").strip().lower() == "unknown"


def _board_type_label_from_schedule_type(schedule_type):
    """Return user-facing board type from PanelScheduleType enum."""
    if _is_switchboard_schedule_type(schedule_type):
        return "Switchboard"
    if _is_data_panel_schedule_type(schedule_type):
        return "Other Panel"
    return "Panelboard"


def _voltage_from_internal(value):
    """Convert Revit internal voltage to volts when possible."""
    if value is None:
        return None
    try:
        return DB.UnitUtils.ConvertFromInternalUnits(float(value), DB.UnitTypeId.Volts)
    except Exception:
        return value


def _first_valid_element(doc, candidates):
    """Return the first valid element resolved from mixed candidate types."""
    for candidate in list(candidates or []):
        try:
            if candidate is None:
                continue
            if isinstance(candidate, DB.Element):
                return candidate
            if isinstance(candidate, DB.ElementId):
                if candidate == DB.ElementId.InvalidElementId:
                    continue
                element = doc.GetElement(candidate)
                if element is not None:
                    return element
            try:
                element_id = revit_helpers.elementid_from_value(int(candidate))
            except Exception:
                continue
            element = doc.GetElement(element_id)
            if element is not None:
                return element
        except Exception:
            continue
    return None


def _panel_sort_mode_from_configuration(panel_configuration):
    """Return normalized slot-order mode from DBE.PanelConfiguration."""
    def _enum_equals(value, target):
        try:
            if value == target:
                return True
        except Exception:
            pass
        try:
            return int(value) == int(target)
        except Exception:
            pass
        return _to_text(value, "").strip().lower() == _to_text(target, "").strip().lower()

    if _enum_equals(panel_configuration, DBE.PanelConfiguration.OneColumn):
        return SORT_MODE_PANELBOARD_ONE_COLUMN
    if _enum_equals(panel_configuration, DBE.PanelConfiguration.TwoColumnsCircuitsDown):
        return SORT_MODE_PANELBOARD_DOWN
    return SORT_MODE_PANELBOARD_ACROSS


def _layout_from_table_data(table_data):
    """Build PanelLayoutInfo from PanelScheduleData."""
    if table_data is None:
        return PanelLayoutInfo(
            schedule_type=PSTYPE_UNKNOWN,
            panel_configuration=DBE.PanelConfiguration.OneColumn,
            sort_mode=SORT_MODE_PANELBOARD_ONE_COLUMN,
            board_type="Panelboard",
            max_slot=0,
        )

    try:
        schedule_type = table_data.ScheduleType
    except Exception:
        schedule_type = PSTYPE_UNKNOWN
    try:
        panel_configuration = table_data.PanelConfiguration
    except Exception:
        panel_configuration = DBE.PanelConfiguration.OneColumn
    try:
        max_slot = int(table_data.NumberOfSlots or 0)
    except Exception:
        max_slot = 0
    try:
        show_slots_from_device = bool(getattr(table_data, "ShowSlotFromDeviceInsteadOfTemplate", False))
    except Exception:
        show_slots_from_device = False

    if _is_switchboard_schedule_type(schedule_type):
        return PanelLayoutInfo(
            schedule_type=schedule_type,
            panel_configuration=DBE.PanelConfiguration.OneColumn,
            sort_mode=SORT_MODE_SWITCHBOARD,
            board_type="Switchboard",
            max_slot=max_slot,
            show_slots_from_device=show_slots_from_device,
        )
    if _is_data_panel_schedule_type(schedule_type):
        return PanelLayoutInfo(
            schedule_type=schedule_type,
            panel_configuration=DBE.PanelConfiguration.OneColumn,
            sort_mode=SORT_MODE_PANELBOARD_ONE_COLUMN,
            board_type="Other Panel",
            max_slot=max_slot,
            show_slots_from_device=show_slots_from_device,
        )
    return PanelLayoutInfo(
        schedule_type=schedule_type,
        panel_configuration=panel_configuration,
        sort_mode=_panel_sort_mode_from_configuration(panel_configuration),
        board_type="Panelboard",
        max_slot=max_slot,
        show_slots_from_device=show_slots_from_device,
    )


def _classify_layout_heuristic(schedule_view):
    """Fallback heuristic for schedule layout when table metadata is unavailable."""
    if schedule_view is None:
        return SORT_MODE_PANELBOARD_ACROSS
    try:
        table = schedule_view.GetTableData()
        body = table.GetSectionData(DB.SectionType.Body)
        if body is None:
            return SORT_MODE_PANELBOARD_ACROSS
    except Exception:
        return SORT_MODE_PANELBOARD_ACROSS

    two_slot_rows = 0
    one_slot_rows = 0
    for row in range(body.NumberOfRows):
        slots = set()
        for col in range(body.NumberOfColumns):
            try:
                slot = schedule_view.GetSlotNumberByCell(row, col)
            except Exception:
                slot = 0
            if slot and slot > 0:
                slots.add(int(slot))
        if len(slots) >= 2:
            two_slot_rows += 1
        elif len(slots) == 1:
            one_slot_rows += 1
    return SORT_MODE_PANELBOARD_ACROSS if two_slot_rows > one_slot_rows else SORT_MODE_SWITCHBOARD


def get_schedule_layout_info(schedule_source):
    """Return layout metadata for PanelScheduleView or PanelScheduleTemplate."""
    if schedule_source is None:
        return _layout_from_table_data(None)
    try:
        table_data = schedule_source.GetTableData()
    except Exception:
        table_data = None
    layout = _layout_from_table_data(table_data)
    if _is_unknown_schedule_type(layout.get("schedule_type")) and isinstance(schedule_source, DBE.PanelScheduleView):
        mode = _classify_layout_heuristic(schedule_source)
        layout["sort_mode"] = mode
        layout["board_type"] = "Switchboard" if mode == SORT_MODE_SWITCHBOARD else "Panelboard"
    return layout


def get_panel_board_type(panel):
    """Return board label from family part type with MEP fallback."""
    part_type = get_panel_family_part_type(panel)
    if part_type in PART_TYPE_MAP:
        return PART_TYPE_MAP.get(part_type, "Panelboard")
    return "Panelboard"


def get_panel_family_part_type(panel):
    """Return family part type integer from panel family definition."""
    return _distribution_equipment_repo().get_family_part_type(panel)


def get_expected_panel_schedule_type(panel):
    """Return expected PanelScheduleType from family part type."""
    expected = _distribution_equipment_repo().expected_panel_schedule_type_for_equipment(panel)
    if expected is None:
        return PSTYPE_UNKNOWN
    return expected


def get_all_panels(doc, require_mep_model=True, exclude_design_options=True, require_supported_part_type=True):
    """Return panel/switchboard equipment from main model."""
    collector = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment)
        .WhereElementIsNotElementType()
    )
    if bool(exclude_design_options):
        collector = collector.WherePasses(_DESIGN_OPTION_MAIN_MODEL_FILTER)

    results = []
    for panel in collector:
        if not isinstance(panel, DB.FamilyInstance):
            continue
        if not bool(require_mep_model):
            results.append(panel)
            continue
        try:
            mep_model = panel.MEPModel
        except Exception:
            mep_model = None
        if mep_model is None:
            continue
        if isinstance(mep_model, DBE.ElectricalEquipment):
            if bool(require_supported_part_type):
                part_type = get_panel_family_part_type(panel)
                if part_type not in _PART_TYPE_TO_SCHEDULE_TYPE:
                    continue
            results.append(panel)
    return results


def _iter_circuit_elements(circuit):
    """Yield elements connected to an electrical system, best-effort."""
    if circuit is None:
        return
    try:
        elements = getattr(circuit, "Elements", None)
        if elements is not None:
            for element in elements:
                if element is not None:
                    yield element
            return
    except Exception:
        pass
    getter = getattr(circuit, "GetCircuitElements", None)
    if getter:
        try:
            for element in getter():
                if element is not None:
                    yield element
        except Exception:
            pass


def get_circuit_fed_panel_ids(circuit, panel_id_set=None):
    """Return panel ids served by a circuit based on connected load elements."""
    candidate_ids = set([int(x) for x in list(panel_id_set or []) if int(x) > 0])
    fed = set()
    for element in _iter_circuit_elements(circuit):
        try:
            elem_id = _idval(getattr(element, "Id", None))
            if elem_id <= 0:
                continue
            if candidate_ids and elem_id not in candidate_ids:
                continue
            category = getattr(element, "Category", None)
            if category is not None and getattr(category, "Id", None) is not None:
                if _idval(category.Id) != int(DB.BuiltInCategory.OST_ElectricalEquipment):
                    continue
            fed.add(int(elem_id))
        except Exception:
            continue
    return sorted(list(fed))


def get_panel_distribution_profile(doc, panel):
    """Return panel distribution profile used for compatibility checks."""
    result = {
        "dist_system_name": None,
        "phase": None,
        "wire_count": None,
        "lg_voltage": None,
        "ll_voltage": None,
    }
    if panel is None:
        return result

    dist_system_id = None
    for bip in (
        DB.BuiltInParameter.RBS_FAMILY_CONTENT_SECONDARY_DISTRIBSYS,
        DB.BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM,
    ):
        try:
            param = panel.get_Parameter(bip)
            if param and param.HasValue:
                dist_system_id = param.AsElementId()
                if dist_system_id and dist_system_id != DB.ElementId.InvalidElementId:
                    break
        except Exception:
            continue

    if not dist_system_id or dist_system_id == DB.ElementId.InvalidElementId:
        return result

    dist_system = doc.GetElement(dist_system_id)
    if dist_system is None:
        return result

    try:
        name_param = dist_system.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if name_param and name_param.HasValue:
            result["dist_system_name"] = name_param.AsString() or "Unnamed Distribution System"
    except Exception:
        pass

    try:
        result["phase"] = getattr(dist_system, "ElectricalPhase", None)
    except Exception:
        result["phase"] = None

    try:
        value = dist_system.NumWires
        if value is not None:
            wires = int(value)
            if wires > 0:
                result["wire_count"] = wires
    except Exception:
        pass

    try:
        lg_voltage_type = getattr(dist_system, "VoltageLineToGround", None)
        if lg_voltage_type is not None:
            lg_param = lg_voltage_type.get_Parameter(DB.BuiltInParameter.RBS_VOLTAGETYPE_VOLTAGE_PARAM)
            if lg_param and lg_param.HasValue:
                result["lg_voltage"] = _voltage_from_internal(lg_param.AsDouble())
    except Exception:
        pass

    try:
        ll_voltage_type = getattr(dist_system, "VoltageLineToLine", None)
        if ll_voltage_type is not None:
            ll_param = ll_voltage_type.get_Parameter(DB.BuiltInParameter.RBS_VOLTAGETYPE_VOLTAGE_PARAM)
            if ll_param and ll_param.HasValue:
                result["ll_voltage"] = _voltage_from_internal(ll_param.AsDouble())
    except Exception:
        pass

    return result


def get_distribution_display_name(profile):
    """Return display-friendly distribution system name."""
    name = (profile or {}).get("dist_system_name")
    if not name:
        return "Unknown Dist. System"
    return _to_text(name, "Unknown Dist. System")


def format_panel_display(panel, profile=None):
    """Return panel label for single-select UI prompts."""
    if panel is None:
        return "Unknown Panel"
    profile = profile or {}
    dist_name = get_distribution_display_name(profile)
    return "{0} - {1} (ID: {2})".format(
        _to_text(getattr(panel, "Name", ""), "Unnamed Panel") or "Unnamed Panel",
        dist_name,
        _to_text(getattr(panel, "Id", ""), ""),
    )


def get_sorted_panels_for_display(doc, panels=None):
    """Return panels sorted by panel name and distribution system name."""
    panel_list = list(panels or get_all_panels(doc))
    prepared = []
    for panel in panel_list:
        profile = get_panel_distribution_profile(doc, panel)
        dist_name = get_distribution_display_name(profile)
        if dist_name == "Unnamed Distribution System":
            continue
        prepared.append((_to_text(getattr(panel, "Name", ""), ""), dist_name, panel))
    prepared.sort(key=lambda x: (x[0], x[1]))
    return [item[2] for item in prepared]


def prompt_for_panel(doc, panels, title="Select Panel", prompt_msg="Choose a panel", multiselect=False):
    """Prompt user to select one or more panels from a sorted panel list."""
    sorted_panels = get_sorted_panels_for_display(doc, panels)
    if not sorted_panels:
        return [] if multiselect else None

    panel_map = {}
    display_items = []
    for panel in sorted_panels:
        profile = get_panel_distribution_profile(doc, panel)
        label = format_panel_display(panel, profile)
        panel_map[label] = panel
        display_items.append(label)

    selected_display = forms.SelectFromList.show(
        display_items,
        title=title,
        prompt=prompt_msg,
        multiselect=bool(multiselect),
    )
    if not selected_display:
        return [] if multiselect else None

    if bool(multiselect):
        if not isinstance(selected_display, list):
            selected_display = [selected_display]
        return [panel_map[x] for x in selected_display if x in panel_map]
    return panel_map.get(selected_display)


def get_panel_schedule_views(doc):
    """Return non-template panel schedule views."""
    views = []
    for view in DB.FilteredElementCollector(doc).OfClass(DBE.PanelScheduleView).ToElements():
        try:
            if bool(getattr(view, "IsTemplate", False)):
                continue
        except Exception:
            pass
        views.append(view)
    return views


def get_panel_schedule_templates(doc):
    """Return panel schedule templates loaded in the document."""
    templates = []
    for template in DB.FilteredElementCollector(doc).OfClass(DBE.PanelScheduleTemplate).ToElements():
        if template is None:
            continue
        templates.append(template)
    return templates


def resolve_schedule_panel(doc, schedule_view, all_panels=None):
    """Best-effort resolution of panel represented by a panel schedule view."""
    if schedule_view is None:
        return None

    candidates = []
    try:
        getter = getattr(schedule_view, "GetPanel", None)
        if getter:
            candidates.append(getter())
    except Exception:
        pass
    try:
        getter = getattr(schedule_view, "GetPanelId", None)
        if getter:
            candidates.append(getter())
    except Exception:
        pass
    try:
        candidates.append(getattr(schedule_view, "PanelId", None))
    except Exception:
        pass
    try:
        candidates.append(getattr(schedule_view, "Panel", None))
    except Exception:
        pass

    panel = _first_valid_element(doc, candidates)
    if panel is not None:
        return panel

    panels = list(all_panels or get_all_panels(doc))
    schedule_name = _to_text(getattr(schedule_view, "Name", ""), "").strip()
    if not schedule_name:
        return None

    exact = [p for p in panels if _to_text(getattr(p, "Name", ""), "").strip() == schedule_name]
    if len(exact) == 1:
        return exact[0]

    lowered_schedule = schedule_name.lower()
    starts = [p for p in panels if lowered_schedule.startswith(_to_text(getattr(p, "Name", ""), "").strip().lower())]
    if len(starts) == 1:
        return starts[0]

    contains = []
    for panel in panels:
        panel_name = _to_text(getattr(panel, "Name", ""), "").strip()
        if not panel_name:
            continue
        lowered_panel = panel_name.lower()
        if lowered_panel in lowered_schedule or lowered_schedule in lowered_panel:
            contains.append(panel)
    if len(contains) == 1:
        return contains[0]
    if contains:
        contains.sort(key=lambda p: len(_to_text(getattr(p, "Name", ""), "").strip()), reverse=True)
        return contains[0]

    return None


def get_panel_schedule_view_for_panel(doc, panel, views=None):
    """Return panel schedule instance view mapped to a panel, or None."""
    if panel is None:
        return None
    panel_id = _idval(getattr(panel, "Id", None))
    if panel_id <= 0:
        return None
    candidates = list(views or get_panel_schedule_views(doc))
    best_view = None
    best_slots = -1
    for view in candidates:
        resolved = resolve_schedule_panel(doc, view, all_panels=[panel])
        if resolved is None or _idval(resolved.Id) != panel_id:
            continue
        try:
            table = view.GetTableData()
            slots = int(getattr(table, "NumberOfSlots", 0) or 0)
        except Exception:
            slots = 0
        if best_view is None or slots > best_slots:
            best_view = view
            best_slots = slots
    return best_view


def map_panel_schedule_views(doc, panels=None, views=None):
    """Return map of panel_id -> best mapped panel schedule view."""
    panel_map = {}
    panel_by_id = {}
    for panel in list(panels or get_all_panels(doc)):
        panel_id = _idval(getattr(panel, "Id", None))
        if panel_id > 0:
            panel_by_id[panel_id] = panel
    if not panel_by_id:
        return panel_map

    for view in list(views or get_panel_schedule_views(doc)):
        panel = resolve_schedule_panel(doc, view, all_panels=panel_by_id.values())
        if panel is None:
            continue
        panel_id = _idval(panel.Id)
        if panel_id <= 0:
            continue
        existing = panel_map.get(panel_id)
        if existing is None:
            panel_map[panel_id] = view
            continue
        try:
            cur_slots = int(existing.GetTableData().NumberOfSlots or 0)
        except Exception:
            cur_slots = 0
        try:
            new_slots = int(view.GetTableData().NumberOfSlots or 0)
        except Exception:
            new_slots = 0
        if new_slots > cur_slots:
            panel_map[panel_id] = view
    return panel_map


def get_panel_schedules_from_selection(doc, selection_elements):
    """Resolve selected schedule elements to PanelScheduleView objects."""
    found = {}
    for element in list(selection_elements or []):
        if element is None:
            continue
        view = None
        if isinstance(element, DBE.PanelScheduleView):
            view = element
        elif isinstance(element, DBE.PanelScheduleSheetInstance):
            try:
                view = doc.GetElement(element.ScheduleId)
            except Exception:
                view = None
        if isinstance(view, DBE.PanelScheduleView):
            found[_idval(view.Id)] = view
    return list(found.values())


def collect_schedules_to_process(doc, active_view=None, selection_elements=None, prompt_if_empty=True):
    """Resolve panel schedules from active view, selection, or optional prompt."""
    if isinstance(active_view, DBE.PanelScheduleView):
        return [active_view]

    normalized = []
    for item in list(selection_elements or []):
        try:
            if isinstance(item, DB.Element):
                normalized.append(item)
            elif isinstance(item, DB.ElementId):
                el = doc.GetElement(item)
                if el is not None:
                    normalized.append(el)
        except Exception:
            continue

    picked = get_panel_schedules_from_selection(doc, normalized)
    if picked:
        return picked

    if not bool(prompt_if_empty):
        return []
    return prompt_pick_panel_schedules(doc, title="Choose panel schedules", multiselect=True)


def prompt_pick_panel_schedules(doc, title="Choose panel schedules", multiselect=True):
    """Prompt user to choose panel schedule views."""
    views = get_panel_schedule_views(doc)
    if not views:
        return []

    class _ScheduleOption(object):
        def __init__(self, view):
            self.view = view

        def __str__(self):
            return _to_text(getattr(self.view, "Name", ""), "Unnamed Schedule") or "Unnamed Schedule"

    options = [_ScheduleOption(v) for v in sorted(views, key=lambda x: _to_text(getattr(x, "Name", ""), ""))]
    picked = forms.SelectFromList.show(options, title=title, multiselect=bool(multiselect))
    if not picked:
        return []
    if not isinstance(picked, list):
        picked = [picked]
    return [x.view for x in picked if hasattr(x, "view")]


def get_cells_by_slot_number(schedule_view, slot, body=None):
    """Return raw API cell coordinates for a slot from PanelScheduleView.GetCellsBySlotNumber."""
    if schedule_view is None:
        return []
    target_slot = int(slot or 0)
    if target_slot <= 0:
        return []
    getter = getattr(schedule_view, "GetCellsBySlotNumber", None)
    if getter is None:
        return []

    def _filter_cells(pairs):
        filtered = []
        for pair in list(pairs or []):
            if not pair or len(pair) < 2:
                continue
            row = int(pair[0])
            col = int(pair[1])
            try:
                slot_no = int(schedule_view.GetSlotNumberByCell(int(row), int(col)) or 0)
            except Exception:
                slot_no = 0
            if int(slot_no) != int(target_slot):
                continue
            filtered.append((int(row), int(col)))
        return sorted(set(filtered))

    # Pattern A: out-ref (most reliable across pythonnet bindings)
    try:
        row_ref = clr.Reference[IList[int]]()
        col_ref = clr.Reference[IList[int]]()
        getter(int(target_slot), row_ref, col_ref)
        row_arr = list(row_ref.Value or [])
        col_arr = list(col_ref.Value or [])
        pair_count = int(min(len(row_arr), len(col_arr)))
        pairs = []
        for idx in range(pair_count):
            pairs.append((int(row_arr[idx]), int(col_arr[idx])))
        filtered = _filter_cells(pairs)
        if filtered:
            return filtered
    except Exception:
        pass

    # Pattern B: tuple-return overload
    try:
        direct = getter(int(target_slot))
    except Exception:
        direct = None
    tuple_pairs = []
    if isinstance(direct, tuple) and len(direct) >= 2:
        try:
            row_arr = list(direct[0] or [])
            col_arr = list(direct[1] or [])
            pair_count = int(min(len(row_arr), len(col_arr)))
            for idx in range(pair_count):
                tuple_pairs.append((int(row_arr[idx]), int(col_arr[idx])))
        except Exception:
            tuple_pairs = []
    if tuple_pairs:
        filtered = _filter_cells(tuple_pairs)
        if filtered:
            return filtered

    # Pattern C: enumerable of row/col pair objects
    pairs = []
    for item in list(direct or []):
        row = None
        col = None
        for pair in (("Row", "Column"), ("RowNumber", "ColumnNumber"), ("Item1", "Item2")):
            try:
                row = int(getattr(item, pair[0]))
                col = int(getattr(item, pair[1]))
                break
            except Exception:
                row = None
                col = None
        if row is None or col is None:
            try:
                row = int(item[0])
                col = int(item[1])
            except Exception:
                continue
        pairs.append((int(row), int(col)))
    if pairs:
        filtered = _filter_cells(pairs)
        if filtered:
            return filtered

    # Pattern D: preallocated lists (last fallback)
    try:
        row_arr = List[int]()
        col_arr = List[int]()
        getter(int(target_slot), row_arr, col_arr)
        pair_count = int(min(len(row_arr), len(col_arr)))
        pairs = []
        for idx in range(pair_count):
            pairs.append((int(row_arr[idx]), int(col_arr[idx])))
        return _filter_cells(pairs)
    except Exception:
        return []


def _build_slot_cell_map(schedule_view, body, max_slot):
    """Return map of slot number to body cells by scanning table once."""
    slot_map = {}
    if schedule_view is None or body is None:
        return slot_map
    max_slot_value = int(max_slot or 0)
    if max_slot_value <= 0:
        return slot_map

    for row in range(body.NumberOfRows):
        for col in range(body.NumberOfColumns):
            try:
                slot = int(schedule_view.GetSlotNumberByCell(row, col) or 0)
            except Exception:
                slot = 0
            if slot <= 0 or slot > max_slot_value:
                continue
            if slot not in slot_map:
                slot_map[slot] = []
            slot_map[slot].append((row, col))

    for slot in list(slot_map.keys()):
        slot_map[slot] = sorted(set(slot_map[slot]))
    return slot_map


def is_slot_grouped(schedule_view, slot, body=None):
    """Return True when a slot contains grouped/merged schedule cells."""
    return get_slot_group_number(schedule_view, slot, body=body) > 0


def get_slot_group_number(schedule_view, slot, body=None):
    """Return slot group number, or 0 when slot is not grouped."""
    if schedule_view is None:
        return 0
    slot_value = int(slot or 0)
    if slot_value <= 0:
        return 0
    cells = get_cells_by_slot_number(schedule_view, slot_value, body=body)
    if not cells:
        return 0

    slot_group_fn = getattr(schedule_view, "IsSlotGrouped", None)
    if slot_group_fn is not None:
        for row, col in cells:
            try:
                # Revit API: IsSlotGrouped(row, col) -> Int32 group number
                group_no = int(slot_group_fn(row, col))
                if group_no > 0:
                    return int(group_no)
            except TypeError:
                # Fallback for versions exposing slot-based overload.
                try:
                    if bool(slot_group_fn(slot_value)):
                        return 1
                except Exception:
                    pass
            except Exception:
                continue

    for row, col in cells:
        for attr in ("IsCellGrouped", "IsCellMerged", "GetCellMerged"):
            try:
                grouped_fn = getattr(schedule_view, attr, None)
                if grouped_fn is not None and bool(grouped_fn(row, col)):
                    return 1
            except Exception:
                continue
        try:
            if body is not None:
                grouped_fn = getattr(body, "IsCellMerged", None)
                if grouped_fn is not None and bool(grouped_fn(row, col)):
                    return 1
        except Exception:
            pass
    return 0


def gather_empty_slot_cells(schedule_view):
    """Return map: slot_number -> [(row, col), ...] for empty body cells."""
    empties = {}
    if schedule_view is None:
        return empties
    try:
        table = schedule_view.GetTableData()
        body = table.GetSectionData(DB.SectionType.Body)
    except Exception:
        return empties
    if body is None:
        return empties

    try:
        max_slot = int(getattr(table, "NumberOfSlots", 0) or 0)
    except Exception:
        max_slot = 0
    if max_slot <= 0:
        return empties

    valid_slots = list(range(1, int(max_slot) + 1))
    try:
        doc = getattr(schedule_view, "Document", None)
        panel = resolve_schedule_panel(doc, schedule_view) if doc is not None else None
        if panel is not None:
            option = _panel_option_from_panel_and_view(doc, panel, schedule_view=schedule_view)
            candidate = list(get_option_valid_slots(option) or [])
            if candidate:
                valid_slots = [int(x) for x in list(candidate or []) if int(x) > 0]
    except Exception:
        pass

    for slot in list(valid_slots or []):
        slot_cells = list(get_cells_by_slot_number(schedule_view, int(slot), body=body) or [])
        if not slot_cells:
            continue
        empty_cells = []
        for row, col in slot_cells:
            try:
                ckt_id = schedule_view.GetCircuitIdByCell(int(row), int(col))
            except Exception:
                ckt_id = DB.ElementId.InvalidElementId
            if ckt_id is None or ckt_id == DB.ElementId.InvalidElementId:
                empty_cells.append((int(row), int(col)))
        if empty_cells:
            empties[int(slot)] = sorted(set(empty_cells))

    return empties


def classify_schedule_layout(schedule_view):
    """Return normalized sort mode for the schedule layout."""
    if schedule_view is None:
        return SORT_MODE_PANELBOARD_ACROSS

    layout = get_schedule_layout_info(schedule_view)
    mode = _to_text(layout.get("sort_mode"), SORT_MODE_PANELBOARD_ACROSS).strip().lower()
    if mode:
        return mode
    return _classify_layout_heuristic(schedule_view)


def get_slot_order(max_slot, sort_mode):
    """Return display/placement slot order for a schedule layout."""
    slot_count = int(max(0, max_slot or 0))
    slots = list(range(1, slot_count + 1))
    mode = _to_text(sort_mode, SORT_MODE_PANELBOARD_ACROSS).strip().lower()
    if mode == "panelboard":
        mode = SORT_MODE_PANELBOARD_ACROSS
    if mode == SORT_MODE_PANELBOARD_ACROSS:
        return sorted(slots, key=lambda x: (x % 2 == 0, x))
    return slots


def get_slot_row_order(max_slot, sort_mode):
    """Return row-wise slot order (top-to-bottom rows) independent of add sequencing."""
    slot_count = int(max(0, max_slot or 0))
    if slot_count <= 0:
        return []
    mode = _to_text(sort_mode, SORT_MODE_PANELBOARD_ACROSS).strip().lower()
    if mode == "panelboard":
        mode = SORT_MODE_PANELBOARD_ACROSS
    if mode == SORT_MODE_PANELBOARD_DOWN:
        left_count = int((slot_count + 1) / 2)
        ordered = []
        for row_slot in range(1, left_count + 1):
            ordered.append(int(row_slot))
            right_slot = int(left_count + row_slot)
            if right_slot <= slot_count:
                ordered.append(int(right_slot))
        return ordered
    return list(range(1, slot_count + 1))


def _int_or_zero(value):
    try:
        return int(round(float(value)))
    except Exception:
        return 0


def _param_int_or_zero(param):
    value = revit_helpers.get_parameter_value(param, default=None)
    return _int_or_zero(value)


def _panel_device_slot_capacity(option):
    """Return equipment-based slot capacity (poles/circuits) for a panel option."""
    panel = (option or {}).get("panel")
    schedule_type = (option or {}).get("schedule_type")
    model = (option or {}).get("equipment_model")
    schedule_slots = _int_or_zero((option or {}).get("max_slot", 0))

    # Data panel capacity is based on the schedule slot count.
    if _is_data_panel_schedule_type(schedule_type):
        if schedule_slots > 0:
            return int(schedule_slots)

    # Prefer the already-built distribution equipment model value.
    if model is not None:
        value = _int_or_zero(getattr(model, "max_poles", 0))
        if value:
            return value

    if panel is not None and schedule_type == PSTYPE_SWITCHBOARD:
        try:
            param = panel.get_Parameter(DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_CIRCUITS)
        except Exception:
            param = None
        value = _param_int_or_zero(param)
        if value:
            return value
    if panel is not None:
        try:
            param = panel.get_Parameter(DB.BuiltInParameter.RBS_ELEC_MAX_POLE_BREAKERS)
        except Exception:
            param = None
        value = _param_int_or_zero(param)
        if value:
            return value
    return 0


def _compute_option_slot_limits(option):
    """Compute schedule/device slot limits and derived valid/invalid slot sets."""
    schedule_slots = _int_or_zero((option or {}).get("max_slot", 0))
    sort_mode = (option or {}).get("sort_mode", SORT_MODE_PANELBOARD_ACROSS)
    show_from_device = bool((option or {}).get("show_slots_from_device", False))
    device_capacity = _int_or_zero((option or {}).get("device_slot_capacity", 0))
    if not device_capacity:
        device_capacity = _panel_device_slot_capacity(option)

    all_slots = list(get_slot_order(schedule_slots, sort_mode) or [])
    valid_slots = list(all_slots)
    invalid_slots = []
    has_excess = False

    if (
        schedule_slots > 0
        and not bool(show_from_device)
        and int(device_capacity) > 0
        and int(schedule_slots) > int(device_capacity)
    ):
        row_order = list(get_slot_row_order(schedule_slots, sort_mode) or [])
        valid_set = set([int(x) for x in list(row_order[:int(device_capacity)]) if int(x) > 0])
        valid_slots = [int(x) for x in list(all_slots or []) if int(x) in valid_set]
        invalid_slots = [int(x) for x in list(all_slots or []) if int(x) not in valid_set]
        has_excess = bool(invalid_slots)

    return {
        "schedule_slot_count": int(schedule_slots),
        "device_slot_capacity": int(device_capacity),
        "show_slots_from_device": bool(show_from_device),
        "all_slots": [int(x) for x in list(all_slots or []) if int(x) > 0],
        "valid_slots": [int(x) for x in list(valid_slots or []) if int(x) > 0],
        "invalid_slots": [int(x) for x in list(invalid_slots or []) if int(x) > 0],
        "usable_slot_count": int(len([x for x in list(valid_slots or []) if int(x) > 0])),
        "has_excess_slots": bool(has_excess),
    }


def get_option_slot_limits(option):
    """Return slot limits map for a panel option."""
    cached = {}
    try:
        cached = dict((option or {}).get("slot_limits") or {})
    except Exception:
        cached = {}
    if cached and "valid_slots" in cached and "all_slots" in cached:
        return cached
    return _compute_option_slot_limits(option)


def get_option_slot_order(option, include_excess=False):
    """Return option slot order, filtered to valid slots by default."""
    limits = get_option_slot_limits(option)
    if bool(include_excess):
        return [int(x) for x in list(limits.get("all_slots") or []) if int(x) > 0]
    return [int(x) for x in list(limits.get("valid_slots") or []) if int(x) > 0]


def get_option_valid_slots(option):
    """Return valid slot list for add/move operations on this option."""
    return list(get_option_slot_order(option, include_excess=False))


def get_option_invalid_slots(option):
    """Return excess template-only slot list that exceeds device capacity."""
    limits = get_option_slot_limits(option)
    return [int(x) for x in list(limits.get("invalid_slots") or []) if int(x) > 0]


def get_option_usable_slot_count(option):
    """Return count of valid usable slots for this option."""
    limits = get_option_slot_limits(option)
    return int(limits.get("usable_slot_count", 0) or 0)


def is_slot_valid_for_option(option, slot):
    """Return True when slot is inside equipment-supported slot range for this option."""
    slot_value = int(slot or 0)
    if slot_value <= 0:
        return False
    valid_slots = set([int(x) for x in list(get_option_valid_slots(option) or []) if int(x) > 0])
    if valid_slots:
        return bool(slot_value in valid_slots)
    max_slot = int((option or {}).get("max_slot", 0) or 0)
    return bool(max_slot > 0 and slot_value <= max_slot)


def get_slot_span_slots_for_option(option, start_slot, pole_count, require_valid=False):
    """Return covered slot span for an option-aware start slot/pole combination."""
    if option is None:
        return []
    slots = get_slot_span_slots(
        start_slot=int(start_slot or 0),
        pole_count=int(max(1, pole_count or 1)),
        max_slot=int(option.get("max_slot", 0) or 0),
        sort_mode=option.get("sort_mode", SORT_MODE_PANELBOARD_ACROSS),
    )
    if not slots:
        return []
    normalized = [int(x) for x in list(slots or []) if int(x) > 0]
    if not bool(require_valid):
        return normalized
    for slot in list(normalized or []):
        if not bool(is_slot_valid_for_option(option, int(slot))):
            return []
    return normalized


def get_slot_display_column(slot, max_slot, sort_mode):
    """Return 1-based display column index for a slot in current layout."""
    slot_value = int(slot or 0)
    if slot_value <= 0:
        return 1
    slot_count = int(max(0, max_slot or 0))
    mode = _to_text(sort_mode, SORT_MODE_PANELBOARD_ACROSS).strip().lower()
    if mode == "panelboard":
        mode = SORT_MODE_PANELBOARD_ACROSS
    if mode == SORT_MODE_PANELBOARD_ACROSS:
        return 1 if int(slot_value % 2) == 1 else 2
    if mode == SORT_MODE_PANELBOARD_DOWN:
        left_count = int((slot_count + 1) / 2)
        return 1 if slot_value <= left_count else 2
    if mode == SORT_MODE_SWITCHBOARD:
        return 1
    return 1


def get_slot_span_slots(start_slot, pole_count, max_slot, sort_mode):
    """Return covered slots for a circuit start slot and pole count."""
    slot_count = int(max(0, max_slot or 0))
    slot_order = get_slot_order(max_slot, sort_mode)
    index_map = {}
    for idx, slot in enumerate(slot_order):
        index_map[int(slot)] = idx

    slot_value = int(start_slot or 0)
    if slot_value <= 0 or slot_value not in index_map:
        return []

    mode = _to_text(sort_mode, SORT_MODE_PANELBOARD_ACROSS).strip().lower()
    if mode == "panelboard":
        mode = SORT_MODE_PANELBOARD_ACROSS

    if mode == SORT_MODE_SWITCHBOARD:
        return [slot_value]

    pole_total = int(max(1, pole_count or 1))
    if pole_total <= 1:
        return [slot_value]

    if mode == SORT_MODE_PANELBOARD_ACROSS:
        parity = int(slot_value % 2)
        same_side_slots = [int(x) for x in list(slot_order) if int(x % 2) == parity]
        if slot_value not in same_side_slots:
            return []
        start_idx = same_side_slots.index(slot_value)
        slots = same_side_slots[start_idx:start_idx + pole_total]
        if len(slots) < pole_total:
            return []
        return [int(x) for x in slots]

    if mode == SORT_MODE_PANELBOARD_DOWN:
        left_count = int((slot_count + 1) / 2)
        if slot_value <= left_count:
            end_slot = slot_value + pole_total - 1
            if end_slot > left_count:
                return []
            return list(range(slot_value, end_slot + 1))
        end_slot = slot_value + pole_total - 1
        if end_slot > slot_count:
            return []
        return list(range(slot_value, end_slot + 1))

    # One-column branch layout behaves like contiguous slots in sequence.
    end_slot = slot_value + pole_total - 1
    if end_slot > slot_count:
        return []
    return list(range(slot_value, end_slot + 1))


def predict_circuit_number(option, start_slot, poles=1):
    """Predict a circuit-number string for a staged row at a slot."""
    slots = get_slot_span_slots(
        start_slot=start_slot,
        pole_count=poles,
        max_slot=option.get("max_slot", 0),
        sort_mode=option.get("sort_mode", SORT_MODE_PANELBOARD_ACROSS),
    )
    if not slots:
        return _to_text(start_slot, "")
    if len(slots) == 1:
        return _to_text(slots[0], "")
    return ",".join([_to_text(x, "") for x in slots])


def get_electrical_settings(doc):
    """Return ElectricalSetting using version-compatible lookups."""
    try:
        return DBE.ElectricalSetting.GetElectricalSettings(doc)
    except Exception:
        pass
    try:
        return doc.Settings.ElectricalSettings
    except Exception:
        return None


def get_default_circuit_rating(electrical_settings):
    """Return default circuit rating from ElectricalSettings."""
    if electrical_settings is None:
        return None
    try:
        return getattr(electrical_settings, "CircuitRating")
    except Exception:
        return None


def _approx_equal(a, b, tol=1e-6):
    """Return True when values are approximately equal."""
    try:
        return abs(float(a) - float(b)) <= float(tol)
    except Exception:
        return False


def _slot_is_locked(schedule_view, row, col):
    """Best-effort slot lock state for a schedule cell."""
    slot_value = 0
    try:
        slot_value = int(schedule_view.GetSlotNumberByCell(int(row), int(col)) or 0)
    except Exception:
        slot_value = 0
    for attr in ("IsSlotLocked", "GetLockSlot", "IsCellLocked"):
        fn = getattr(schedule_view, attr, None)
        if fn:
            for args in (
                (int(row), int(col)),
                (int(slot_value),) if int(slot_value) > 0 else None,
            ):
                if args is None:
                    continue
                try:
                    return bool(fn(*args))
                except Exception:
                    continue
    return False


def get_element_edited_by(doc, element):
    """Return worksharing owner/user for an element, or empty string."""
    if element is None:
        return ""
    try:
        param = element.get_Parameter(DB.BuiltInParameter.EDITED_BY)
        value = revit_helpers.get_parameter_value(param, default=None)
        if value:
            return _to_text(value, "").strip()
    except Exception:
        pass
    try:
        info = DB.WorksharingUtils.GetWorksharingTooltipInfo(doc, element.Id)
        owner = _to_text(getattr(info, "Owner", ""), "").strip()
        if owner:
            return owner
    except Exception:
        pass
    return ""


def is_removable_spare(schedule_view, row, col, circuit, electrical_settings=None, default_rating=None):
    """Return True when a spare row satisfies removable rules."""
    if schedule_view is None or circuit is None:
        return False
    if _slot_is_locked(schedule_view, row, col):
        return False
    try:
        if not bool(schedule_view.IsSpare(row, col)):
            return False
    except Exception:
        return False

    load_name = _to_text(getattr(circuit, "LoadName", ""), "").strip().lower()
    if load_name != "spare":
        return False

    try:
        notes_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM)
        notes = revit_helpers.get_parameter_value(notes_param, default=None)
        if notes:
            return False
    except Exception:
        return False

    try:
        apparent = float(getattr(circuit, "ApparentLoad", 0.0) or 0.0)
        if abs(apparent) > 1e-6:
            return False
    except Exception:
        return False

    rating = None
    try:
        rating = float(getattr(circuit, "Rating", None))
    except Exception:
        rating = None
    if rating is None:
        return False

    if default_rating is None and electrical_settings is not None:
        default_rating = get_default_circuit_rating(electrical_settings)
    if default_rating is None:
        return False
    return _approx_equal(rating, default_rating)


def is_removable_space(schedule_view, row, col, circuit):
    """Return True when a space row satisfies removable rules."""
    if schedule_view is None or circuit is None:
        return False
    if _slot_is_locked(schedule_view, row, col):
        return False
    is_space_flag = False
    try:
        is_space_flag = bool(schedule_view.IsSpace(row, col))
    except Exception:
        pass
    ctype = getattr(circuit, "CircuitType", None)
    type_is_space = bool(ctype == DBE.CircuitType.Space)
    load_name = _to_text(getattr(circuit, "LoadName", ""), "").strip().lower()
    if not ((is_space_flag or type_is_space) and load_name == "space"):
        return False
    try:
        notes_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM)
        notes = revit_helpers.get_parameter_value(notes_param, default=None)
        if notes:
            return False
    except Exception:
        return False
    return True


def _panel_option_from_panel_and_view(doc, panel, schedule_view=None):
    """Build a normalized panel option from equipment and optional schedule view."""
    panel_id = _idval(getattr(panel, "Id", None))
    equipment_model = _distribution_equipment_repo().build_distribution_equipment(
        doc, panel, schedule_view=schedule_view
    )
    part_type = get_panel_family_part_type(panel)
    part_type_name = PART_TYPE_MAP.get(part_type, "Unknown")
    if equipment_model is not None and _to_text(getattr(equipment_model, "equipment_type", ""), ""):
        part_type_name = _to_text(getattr(equipment_model, "equipment_type", ""), part_type_name)
    family_board_type = part_type_name if part_type_name != "Unknown" else get_panel_board_type(panel)
    profile = get_panel_distribution_profile(doc, panel)
    dist_name = get_distribution_display_name(profile)

    if isinstance(schedule_view, DBE.PanelScheduleView):
        layout = get_schedule_layout_info(schedule_view)
        if family_board_type and family_board_type != "Unknown":
            layout["board_type"] = family_board_type
        try:
            schedule_type = layout.get("schedule_type")
        except Exception:
            schedule_type = None
        if _is_unknown_schedule_type(schedule_type):
            expected_type = get_expected_panel_schedule_type(panel)
            if not _is_unknown_schedule_type(expected_type):
                layout["schedule_type"] = expected_type
                layout["schedule_type_name"] = _to_text(expected_type, "Unknown")
                schedule_type = expected_type
        if _is_switchboard_schedule_type(schedule_type):
            layout["sort_mode"] = SORT_MODE_SWITCHBOARD
            layout["panel_configuration"] = DBE.PanelConfiguration.OneColumn
        elif _is_data_panel_schedule_type(schedule_type):
            layout["sort_mode"] = SORT_MODE_PANELBOARD_ONE_COLUMN
            layout["panel_configuration"] = DBE.PanelConfiguration.OneColumn
        else:
            panel_configuration = layout.get("panel_configuration")
            if panel_configuration is None and equipment_model is not None:
                try:
                    panel_configuration = getattr(equipment_model, "panel_configuration", None)
                except Exception:
                    panel_configuration = None
            layout["panel_configuration"] = panel_configuration
            layout["sort_mode"] = _panel_sort_mode_from_configuration(panel_configuration)
        schedule_id = _idval(schedule_view.Id)
        schedule_name = _to_text(getattr(schedule_view, "Name", ""), "Unnamed Schedule") or "Unnamed Schedule"
    else:
        schedule_type = get_expected_panel_schedule_type(panel)
        panel_configuration = None
        if equipment_model is not None:
            try:
                model_cfg = getattr(equipment_model, "panel_configuration", None)
                if model_cfg is not None:
                    panel_configuration = model_cfg
            except Exception:
                pass
        if panel_configuration is None:
            if _is_switchboard_schedule_type(schedule_type) or _is_data_panel_schedule_type(schedule_type):
                panel_configuration = DBE.PanelConfiguration.OneColumn
            else:
                panel_configuration = DBE.PanelConfiguration.OneColumn
        if _is_switchboard_schedule_type(schedule_type):
            sort_mode = SORT_MODE_SWITCHBOARD
        else:
            sort_mode = _panel_sort_mode_from_configuration(panel_configuration)
        layout = PanelLayoutInfo(
            schedule_type=schedule_type,
            panel_configuration=panel_configuration,
            sort_mode=sort_mode,
            board_type=family_board_type,
            max_slot=0,
        )
        schedule_id = 0
        schedule_name = ""

    board_type = _to_text(layout.get("board_type"), "Panelboard")
    schedule_slots = int(layout.get("max_slot", 0) or 0)
    show_slots_from_device = bool(layout.get("show_slots_from_device", False))
    limits_probe = {
        "panel": panel,
        "schedule_type": layout.get("schedule_type"),
        "equipment_model": equipment_model,
        "max_slot": int(schedule_slots),
        "sort_mode": layout.get("sort_mode", SORT_MODE_PANELBOARD_ACROSS),
        "show_slots_from_device": bool(show_slots_from_device),
    }
    device_slot_capacity = int(_panel_device_slot_capacity(limits_probe) or 0)
    limits_probe["device_slot_capacity"] = int(device_slot_capacity)
    slot_limits = _compute_option_slot_limits(limits_probe)

    option = PanelEquipmentOption(
        {
            "panel": panel,
            "schedule_view": schedule_view,
            "panel_id": panel_id,
            "schedule_id": int(schedule_id or 0),
            "panel_name": _to_text(getattr(panel, "Name", ""), "Unnamed Panel") or "Unnamed Panel",
            "schedule_name": schedule_name,
            "has_schedule": bool(isinstance(schedule_view, DBE.PanelScheduleView)),
            "missing_schedule": not bool(isinstance(schedule_view, DBE.PanelScheduleView)),
            "equipment_model": equipment_model,
            "dist_system_name": dist_name,
            "profile": profile,
            "sort_mode": layout.get("sort_mode", SORT_MODE_PANELBOARD_ACROSS),
            "board_type": board_type,
            "max_slot": int(schedule_slots),
            "show_slots_from_device": bool(show_slots_from_device),
            "device_slot_capacity": int(slot_limits.get("device_slot_capacity", 0) or 0),
            "usable_slot_count": int(slot_limits.get("usable_slot_count", 0) or 0),
            "has_excess_slots": bool(slot_limits.get("has_excess_slots", False)),
            "valid_slots": [int(x) for x in list(slot_limits.get("valid_slots") or []) if int(x) > 0],
            "invalid_slots": [int(x) for x in list(slot_limits.get("invalid_slots") or []) if int(x) > 0],
            "slot_limits": dict(slot_limits),
            "schedule_type": layout.get("schedule_type"),
            "schedule_type_name": _to_text(layout.get("schedule_type_name", ""), ""),
            "part_type": part_type,
            "part_type_name": part_type_name,
            "panel_configuration": layout.get("panel_configuration"),
            "panel_configuration_name": _to_text(layout.get("panel_configuration_name", ""), ""),
            "layout_info": layout,
            "display_name": "{0} | {1} | {2}".format(
                _to_text(getattr(panel, "Name", ""), "Unnamed Panel") or "Unnamed Panel",
                board_type,
                dist_name,
            ),
        }
    )
    return option


def collect_panel_equipment_options(doc, panels=None, include_without_schedule=True):
    """Return options for all panel/switchboard equipment, with schedule metadata when present."""
    panel_list = list(panels or get_all_panels(doc, require_mep_model=True, exclude_design_options=True))
    panel_list = [p for p in panel_list if p is not None]
    panel_list.sort(key=lambda x: _to_text(getattr(x, "Name", ""), ""))

    mapped_views = map_panel_schedule_views(doc, panels=panel_list)
    options = []
    for panel in panel_list:
        panel_id = _idval(getattr(panel, "Id", None))
        if panel_id <= 0:
            continue
        view = mapped_views.get(panel_id)
        if view is None and not bool(include_without_schedule):
            continue
        options.append(_panel_option_from_panel_and_view(doc, panel, schedule_view=view))
    return options


def collect_panel_schedule_options(doc, panels=None, unique_by_panel=False):
    """Return panel options with mapped schedule views for legacy callers."""
    options = collect_panel_equipment_options(
        doc,
        panels=panels,
        include_without_schedule=False,
    )
    if not bool(unique_by_panel):
        return options
    deduped = {}
    for option in options:
        panel_id = int(option.get("panel_id", 0) or 0)
        if panel_id <= 0:
            continue
        existing = deduped.get(panel_id)
        if existing is None:
            deduped[panel_id] = option
            continue
        if int(option.get("max_slot", 0) or 0) > int(existing.get("max_slot", 0) or 0):
            deduped[panel_id] = option
    return list(deduped.values())


def attach_schedule_to_option(doc, option, schedule_view):
    """Update an existing panel option with a newly created schedule view."""
    if option is None:
        return None
    panel = option.get("panel")
    if panel is None:
        return option
    refreshed = _panel_option_from_panel_and_view(doc, panel, schedule_view=schedule_view)
    option.clear()
    option.update(refreshed)
    return option


def get_compatible_panel_schedule_templates(doc, panel, probe_assignability=True):
    """Return compatible panel schedule templates for a specific panel equipment instance."""
    if panel is None:
        return []
    expected_type = get_expected_panel_schedule_type(panel)
    templates = list(get_panel_schedule_templates(doc))
    templates.sort(key=lambda x: _to_text(getattr(x, "Name", ""), ""))

    can_probe = bool(probe_assignability)
    probe_tx = None
    if can_probe:
        try:
            probe_tx = DB.Transaction(doc, "Probe Panel Schedule Templates")
            probe_tx.Start()
        except Exception:
            can_probe = False
            probe_tx = None

    options = []
    try:
        for template in templates:
            if template is None:
                continue
            try:
                schedule_type = template.GetPanelScheduleType()
            except Exception:
                schedule_type = PSTYPE_UNKNOWN
            if not _is_unknown_schedule_type(expected_type) and not _is_unknown_schedule_type(schedule_type):
                if schedule_type != expected_type:
                    continue

            layout = get_schedule_layout_info(template)
            is_compatible = True
            if can_probe and probe_tx is not None:
                sub = DB.SubTransaction(doc)
                try:
                    sub.Start()
                    probe_view = DBE.PanelScheduleView.CreateInstanceView(doc, template.Id, panel.Id)
                    if probe_view is None:
                        is_compatible = False
                    else:
                        layout = get_schedule_layout_info(probe_view)
                except Exception:
                    is_compatible = False
                finally:
                    try:
                        if sub.GetStatus() == DB.TransactionStatus.Started:
                            sub.RollBack()
                    except Exception:
                        pass

            if not is_compatible:
                continue

            board_type = _to_text(layout.get("board_type", ""), "") or _board_type_label_from_schedule_type(schedule_type)
            panel_config_name = _to_text(layout.get("panel_configuration_name", ""), "Unknown")
            option = PanelScheduleTemplateOption(
                {
                    "template": template,
                    "template_id": _idval(template.Id),
                    "template_name": _to_text(getattr(template, "Name", ""), "Unnamed Template") or "Unnamed Template",
                    "schedule_type": schedule_type,
                    "schedule_type_name": _to_text(schedule_type, "Unknown"),
                    "panel_configuration": layout.get("panel_configuration"),
                    "panel_configuration_name": panel_config_name,
                    "sort_mode": layout.get("sort_mode", SORT_MODE_PANELBOARD_ACROSS),
                    "board_type": board_type,
                    "display_name": "{0} | {1} | {2}".format(
                        _to_text(getattr(template, "Name", ""), "Unnamed Template") or "Unnamed Template",
                        board_type,
                        panel_config_name,
                    ),
                }
            )
            options.append(option)
    finally:
        if probe_tx is not None:
            try:
                if probe_tx.GetStatus() == DB.TransactionStatus.Started:
                    probe_tx.RollBack()
            except Exception:
                pass

    return options


def create_panel_schedule_instance_view(doc, template_id, panel_id):
    """Create a panel schedule instance view from template/panel ids."""
    tid = template_id if isinstance(template_id, DB.ElementId) else revit_helpers.elementid_from_value(template_id)
    pid = panel_id if isinstance(panel_id, DB.ElementId) else revit_helpers.elementid_from_value(panel_id)
    return DBE.PanelScheduleView.CreateInstanceView(doc, tid, pid)


def get_circuit_voltage_poles(circuit):
    """Return (voltage_in_volts, poles) for a circuit."""
    voltage = None
    poles = None
    if circuit is None:
        return voltage, poles

    try:
        voltage_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_VOLTAGE)
        if voltage_param and voltage_param.HasValue:
            voltage = _voltage_from_internal(voltage_param.AsDouble())
    except Exception:
        voltage = None

    try:
        poles_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)
        if poles_param and poles_param.HasValue:
            poles = int(poles_param.AsInteger())
    except Exception:
        poles = None

    return voltage, poles


def get_compatible_panels_for_circuit(circuit, panels, doc, tolerance=1.0):
    """Return panels with matching voltage profile for a circuit."""
    voltage, poles = get_circuit_voltage_poles(circuit)
    if voltage is None or poles is None:
        return []
    poles = int(max(1, poles))

    compatible = []
    for panel in list(panels or []):
        profile = get_panel_distribution_profile(doc, panel)
        if poles <= 1:
            target_voltage = profile.get("lg_voltage") or profile.get("ll_voltage")
        else:
            target_voltage = profile.get("ll_voltage") or profile.get("lg_voltage")
        if target_voltage is None:
            continue
        try:
            if abs(float(target_voltage) - float(voltage)) <= float(tolerance):
                compatible.append(panel)
        except Exception:
            continue
    return compatible


def get_circuit_start_slot(circuit):
    """Return circuit start slot number or 0 when unavailable."""
    try:
        param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_START_SLOT)
        if param and param.HasValue:
            return int(param.AsInteger())
    except Exception:
        pass
    return 0


def _kind_from_circuit(circuit):
    """Return normalized kind for circuit/special rows."""
    ctype = getattr(circuit, "CircuitType", None)
    if ctype == DBE.CircuitType.Space:
        return "space"
    if ctype == DBE.CircuitType.Spare:
        return "spare"
    return "circuit"


def _build_circuit_row(option, circuit, doc=None, panel_id_set=None):
    """Build a normalized row dictionary for a circuit-like item."""
    slot = get_circuit_start_slot(circuit)
    voltage, poles = get_circuit_voltage_poles(circuit)
    kind = _kind_from_circuit(circuit)
    poles_value = int(max(1, poles or 1))

    covered_slots = get_slot_span_slots(
        start_slot=slot,
        pole_count=poles_value,
        max_slot=option.get("max_slot", 0),
        sort_mode=option.get("sort_mode", SORT_MODE_PANELBOARD_ACROSS),
    )
    if not covered_slots:
        covered_slots = [int(slot)] if int(slot or 0) > 0 else []

    load_name = _to_text(getattr(circuit, "LoadName", ""), "").strip()
    circuit_number = _to_text(getattr(circuit, "CircuitNumber", ""), "").strip()
    if not circuit_number:
        circuit_number = predict_circuit_number(option, slot, poles=poles_value)
    schedule_notes_text = ""
    try:
        notes_param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM)
        schedule_notes_text = _to_text(
            revit_helpers.get_parameter_value(notes_param, default=""),
            "",
        ).strip()
    except Exception:
        schedule_notes_text = ""
    voltage_text = "-"
    if voltage is not None:
        try:
            voltage_text = "{0:.0f}V".format(float(voltage))
        except Exception:
            voltage_text = _to_text(voltage, "-")

    rating_value = None
    rating_text = "-"
    try:
        rating_value = float(getattr(circuit, "Rating", None))
        rating_text = "{0:.0f}A".format(rating_value)
    except Exception:
        rating_value = None
        rating_text = "-"

    start_slot = int(covered_slots[0]) if covered_slots else int(slot or 0)
    row_key = "panel:{0}|slot:{1}|ckt:{2}".format(
        option.get("panel_id", 0),
        start_slot,
        _idval(circuit.Id),
    )
    edited_by = ""
    if doc is not None:
        edited_by = get_element_edited_by(doc, circuit)
    fed_panel_ids = []
    if kind == "circuit":
        fed_panel_ids = get_circuit_fed_panel_ids(circuit, panel_id_set=panel_id_set)
    return {
        "row_key": row_key,
        "panel_id": option.get("panel_id", 0),
        "panel_name": option.get("panel_name", ""),
        "slot": start_slot,
        "span": int(max(1, len(covered_slots))),
        "covered_slots": [int(x) for x in covered_slots],
        "kind": kind,
        "is_regular_circuit": (kind == "circuit"),
        "circuit": circuit,
        "circuit_id": _idval(circuit.Id),
        "circuit_number": circuit_number,
        "load_name": load_name,
        "schedule_notes_text": schedule_notes_text,
        "poles": poles_value,
        "voltage": voltage,
        "voltage_text": voltage_text,
        "rating": rating_value,
        "rating_text": rating_text,
        "is_slot_grouped": False,
        "slot_group_number": 0,
        "slot_cells": [],
        "is_spare_removable": False,
        "is_space_removable": False,
        "is_slot_locked": False,
        "edited_by": _to_text(edited_by, ""),
        "is_editable": True,
        "is_valid_slot": True,
        "is_excess_slot": False,
        "fed_panel_ids": list(fed_panel_ids),
    }


def build_empty_row(option, slot, metadata=None):
    """Build a normalized empty slot row for planning UIs."""
    slot_value = int(slot or 0)
    meta = dict(metadata or {})
    is_valid_slot = bool(is_slot_valid_for_option(option, slot_value))
    return {
        "row_key": "panel:{0}|slot:{1}|empty".format(option.get("panel_id", 0), slot_value),
        "panel_id": option.get("panel_id", 0),
        "panel_name": option.get("panel_name", ""),
        "slot": slot_value,
        "span": 1,
        "covered_slots": [slot_value],
        "kind": "empty",
        "is_regular_circuit": False,
        "circuit": None,
        "circuit_id": 0,
        "circuit_number": "",
        "load_name": "",
        "schedule_notes_text": "",
        "poles": 1,
        "voltage": None,
        "voltage_text": "-",
        "rating": None,
        "rating_text": "-",
        "is_slot_grouped": bool(meta.get("is_grouped", False)),
        "slot_group_number": int(meta.get("group_number", 0) or 0),
        "slot_cells": list(meta.get("cells", []) or []),
        "is_spare_removable": False,
        "is_space_removable": False,
        "is_slot_locked": bool(meta.get("is_slot_locked", False)),
        "edited_by": "",
        "is_editable": bool(is_valid_slot),
        "is_valid_slot": bool(is_valid_slot),
        "is_excess_slot": bool(not is_valid_slot),
        "fed_panel_ids": [],
    }


def get_row_covered_slots(row, option=None):
    """Return covered slot list for a row, deriving from poles when needed."""
    row_data = dict(row or {})
    covered = [int(x) for x in list(row_data.get("covered_slots") or []) if int(x) > 0]
    if covered:
        return covered
    slot_value = int(row_data.get("slot", 0) or 0)
    if slot_value <= 0:
        return []
    poles = int(max(1, row_data.get("poles", 1) or 1))
    if option is None:
        return [slot_value]
    return get_slot_span_slots(
        start_slot=slot_value,
        pole_count=poles,
        max_slot=option.get("max_slot", 0),
        sort_mode=option.get("sort_mode", SORT_MODE_PANELBOARD_ACROSS),
    ) or [slot_value]


def _collect_slot_metadata(doc, option):
    """Return slot metadata map for a schedule option."""
    view = option.get("schedule_view")
    if view is None:
        return {}
    metadata = {}
    try:
        table = view.GetTableData()
        body = table.GetSectionData(DB.SectionType.Body)
        max_slot = int(getattr(table, "NumberOfSlots", 0) or 0)
    except Exception:
        return metadata
    if body is None or max_slot <= 0:
        return metadata

    slot_cells_map = _build_slot_cell_map(view, body, max_slot)
    if not slot_cells_map:
        return metadata

    electrical_settings = get_electrical_settings(doc)
    default_rating = get_default_circuit_rating(electrical_settings)
    slot_group_fn = getattr(view, "IsSlotGrouped", None)
    edited_by_cache = {}

    for slot in range(1, max_slot + 1):
        cells = list(slot_cells_map.get(int(slot), []) or [])
        if not cells:
            continue
        row, col = cells[0]
        circuit = None
        circuit_id = DB.ElementId.InvalidElementId
        try:
            circuit_id = view.GetCircuitIdByCell(row, col)
            if circuit_id and circuit_id != DB.ElementId.InvalidElementId:
                circuit = doc.GetElement(circuit_id)
        except Exception:
            circuit = None

        is_spare = False
        is_space = False
        is_locked = _slot_is_locked(view, row, col)
        try:
            is_spare = bool(view.IsSpare(row, col))
        except Exception:
            pass
        try:
            is_space = bool(view.IsSpace(row, col))
        except Exception:
            pass

        kind = "empty"
        if isinstance(circuit, DBE.ElectricalSystem):
            kind = _kind_from_circuit(circuit)
            if kind == "spare":
                is_spare = True
            if kind == "space":
                is_space = True

        is_spare_removable = False
        is_space_removable = False
        group_number = 0
        if slot_group_fn is not None:
            for c_row, c_col in cells:
                try:
                    group_number = int(slot_group_fn(c_row, c_col))
                    if group_number > 0:
                        break
                except TypeError:
                    try:
                        if bool(slot_group_fn(int(slot))):
                            group_number = 1
                            break
                    except Exception:
                        pass
                except Exception:
                    continue
        if group_number <= 0:
            group_number = int(get_slot_group_number(view, slot, body=body) or 0)
        if isinstance(circuit, DBE.ElectricalSystem):
            if is_spare:
                is_spare_removable = is_removable_spare(
                    view,
                    row,
                    col,
                    circuit,
                    electrical_settings=electrical_settings,
                    default_rating=default_rating,
                )
            if is_space:
                is_space_removable = is_removable_space(view, row, col, circuit)
        edited_by = ""
        circuit_id_val = _idval(circuit_id) if circuit_id and circuit_id != DB.ElementId.InvalidElementId else 0
        if circuit_id_val > 0:
            if circuit_id_val in edited_by_cache:
                edited_by = edited_by_cache[circuit_id_val]
            elif isinstance(circuit, DBE.ElectricalSystem):
                edited_by = get_element_edited_by(doc, circuit)
                edited_by_cache[circuit_id_val] = edited_by

        metadata[slot] = {
            "slot": int(slot),
            "cells": list(cells),
            "group_number": int(group_number),
            "is_grouped": bool(group_number > 0),
            "is_slot_locked": bool(is_locked),
            "is_spare": bool(is_spare),
            "is_space": bool(is_space),
            "is_spare_removable": bool(is_spare_removable),
            "is_space_removable": bool(is_space_removable),
            "circuit": circuit,
            "circuit_id": int(circuit_id_val),
            "kind": kind,
            "edited_by": edited_by,
        }
    return metadata


def build_panel_rows(doc, option, panel_id_set=None, all_circuits=None):
    """Build normalized panel rows for UI planning (no transactions)."""
    panel = option.get("panel")
    max_slot = int(option.get("max_slot") or 0)
    sort_mode = option.get("sort_mode") or SORT_MODE_PANELBOARD_ACROSS
    slot_order = list(get_option_slot_order(option, include_excess=True) or [])
    if not slot_order and max_slot > 0:
        slot_order = get_slot_order(max_slot, sort_mode)
    slot_set = set(slot_order)
    valid_slot_set = set([int(x) for x in list(get_option_valid_slots(option) or []) if int(x) > 0])
    if not valid_slot_set and max_slot > 0:
        valid_slot_set = set(range(1, int(max_slot) + 1))
    metadata_map = _collect_slot_metadata(doc, option)
    known_panel_ids = set()
    for candidate_id in list(panel_id_set or []):
        try:
            value = int(candidate_id or 0)
            if value > 0:
                known_panel_ids.add(value)
        except Exception:
            continue
    if not known_panel_ids:
        for candidate_panel in get_all_panels(doc):
            try:
                known_panel_ids.add(int(_idval(candidate_panel.Id)))
            except Exception:
                continue

    circuit_by_slot = {}
    circuits = list(all_circuits) if all_circuits is not None else list(
        DB.FilteredElementCollector(doc).OfClass(DBE.ElectricalSystem).WhereElementIsNotElementType().ToElements()
    )
    for circuit in circuits:
        try:
            base = getattr(circuit, "BaseEquipment", None)
            if base is None or panel is None or _idval(base.Id) != option.get("panel_id", 0):
                continue
        except Exception:
            continue

        slot = get_circuit_start_slot(circuit)
        if slot <= 0:
            continue
        row = _build_circuit_row(option, circuit, doc=doc, panel_id_set=known_panel_ids)
        circuit_by_slot[int(slot)] = row

    rows = []
    consumed = set()
    for slot in slot_order:
        slot_value = int(slot)
        if slot_value in consumed:
            continue

        row = circuit_by_slot.get(slot_value)
        meta = metadata_map.get(slot_value, {})
        if row is None:
            empty = build_empty_row(option, slot_value, metadata=meta)
            rows.append(empty)
            consumed.add(slot_value)
            continue

        covered_slots = [x for x in get_row_covered_slots(row, option=option) if x in slot_set]
        if not covered_slots:
            covered_slots = [slot_value]
        row["slot"] = int(covered_slots[0])
        row["covered_slots"] = list(covered_slots)
        row["span"] = int(max(1, len(covered_slots)))
        row["circuit_number"] = row.get("circuit_number") or predict_circuit_number(
            option,
            row["slot"],
            poles=row.get("poles", 1),
        )
        row["is_slot_grouped"] = bool(meta.get("is_grouped", False))
        row["slot_group_number"] = int(meta.get("group_number", 0) or 0)
        row["slot_cells"] = list(meta.get("cells", []) or [])
        row["is_spare_removable"] = bool(meta.get("is_spare_removable", False))
        row["is_space_removable"] = bool(meta.get("is_space_removable", False))
        row["is_slot_locked"] = bool(meta.get("is_slot_locked", False))
        row["edited_by"] = _to_text(meta.get("edited_by", row.get("edited_by", "")), "")
        row_is_valid = bool(covered_slots) and all(int(x) in valid_slot_set for x in list(covered_slots or []))
        row["is_valid_slot"] = bool(row_is_valid)
        row["is_excess_slot"] = bool(not row_is_valid)
        row["is_editable"] = bool(row_is_valid)
        rows.append(row)
        for covered in covered_slots:
            consumed.add(int(covered))

    return rows


def evaluate_transferability(row, target_option, tolerance=1.0):
    """Return (is_transferable, reason) for moving a circuit row to target panel."""
    if not row:
        return False, "Invalid row."
    if not bool(row.get("is_valid_slot", True)):
        return False, "Slot exceeds equipment-supported capacity."
    kind = _to_text(row.get("kind", ""), "").strip().lower()
    if kind in ("spare", "space"):
        return True, ""
    if not bool(row.get("is_regular_circuit", False)):
        return False, "Only circuits/spares/spaces can be moved."
    target_panel_id = int((target_option or {}).get("panel_id", 0) or 0)
    if target_panel_id > 0:
        fed_ids = set([int(x) for x in list(row.get("fed_panel_ids") or []) if int(x) > 0])
        if target_panel_id in fed_ids:
            return False, "Circuit feeds the target panel/switchboard."

    target_profile = (target_option or {}).get("profile") or {}
    poles = int(row.get("poles") or 1)
    voltage = row.get("voltage")
    if voltage is None:
        return False, "Circuit voltage unavailable."
    target_phase = target_profile.get("phase")
    target_is_single_phase = False
    try:
        if target_phase == DBE.ElectricalPhase.SinglePhase:
            target_is_single_phase = True
    except Exception:
        pass
    if target_is_single_phase and poles >= 3:
        return False, "Target panel is single-phase and cannot accept 3-pole circuits."

    target_model = (target_option or {}).get("equipment_model")
    branch_options = []
    if target_model is not None:
        try:
            branch_options = list(getattr(target_model, "branch_circuit_options", None) or [])
        except Exception:
            branch_options = []
        if not branch_options and isinstance(target_model, dict):
            try:
                branch_options = list(target_model.get("branch_circuit_options") or [])
            except Exception:
                branch_options = []
    if branch_options:
        allowed_poles = set()
        candidate_voltages = []
        for item in list(branch_options or []):
            try:
                option_poles = int(item.get("poles", 0) or 0)
            except Exception:
                option_poles = 0
            if option_poles <= 0:
                continue
            allowed_poles.add(option_poles)
            if option_poles == poles:
                option_voltage = item.get("voltage")
                if option_voltage is not None:
                    candidate_voltages.append(option_voltage)
        if allowed_poles and poles not in allowed_poles:
            return False, "Target panel does not support {0}-pole circuits.".format(int(poles))
        if candidate_voltages:
            if not any(abs(float(v) - float(voltage)) <= float(tolerance) for v in candidate_voltages):
                return False, "Voltage/pole combination is not supported by target panel."

    if poles <= 1:
        target_voltage = target_profile.get("lg_voltage") or target_profile.get("ll_voltage")
        if target_voltage is None:
            return False, "Target panel has no compatible L-G voltage."
        if abs(float(target_voltage) - float(voltage)) > float(tolerance):
            return False, "Voltage mismatch (1-pole requires L-G match)."
        return True, ""

    target_voltage = target_profile.get("ll_voltage") or target_profile.get("lg_voltage")
    if target_voltage is None:
        return False, "Target panel has no compatible L-L voltage."
    if abs(float(target_voltage) - float(voltage)) > float(tolerance):
        return False, "Voltage mismatch ({0}-pole requires L-L match).".format(poles)
    return True, ""
