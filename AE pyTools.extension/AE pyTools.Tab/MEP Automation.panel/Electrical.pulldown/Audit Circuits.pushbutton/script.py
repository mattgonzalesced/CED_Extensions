# -*- coding: utf-8 -*-
__title__ = "Audit Circuits"

import os
import re
import sys
import time
from collections import OrderedDict

from pyrevit import revit, script, forms, DB
from pyrevit.revit.db import query

try:
    import clr
except Exception:
    clr = None

if clr is not None:
    try:
        clr.AddReference("System.Data")
    except Exception:
        pass

try:
    from System.Data import DataTable
except Exception:
    DataTable = None

try:
    from System.ComponentModel import ListSortDirection
except Exception:
    ListSortDirection = None

logger = script.get_logger()
output = script.get_output()

CLIENT_CHOICES = OrderedDict([
    ("Planet Fitness", "planet_fitness"),
    ("HEB", "heb"),
])

SCOPE_CHOICES = OrderedDict([
    ("All eligible elements", False),
    ("Selected elements only", True),
])

EXCLUDED_CATEGORY_IDS = {
    DB.ElementId(DB.BuiltInCategory.OST_LightingDevices).IntegerValue,
    DB.ElementId(DB.BuiltInCategory.OST_LightingFixtures).IntegerValue,
}

OPTION_FILTER = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)


def _collect_by_category(doc, category):
    return (
        DB.FilteredElementCollector(doc)
        .OfCategory(category)
        .WhereElementIsNotElementType()
        .WherePasses(OPTION_FILTER)
    )


def _get_all_panels(doc):
    return _collect_by_category(doc, DB.BuiltInCategory.OST_ElectricalEquipment).ToElements()


def _get_all_elec_fixtures(doc):
    return _collect_by_category(doc, DB.BuiltInCategory.OST_ElectricalFixtures).ToElements()


def _get_all_data_devices(doc):
    return _collect_by_category(doc, DB.BuiltInCategory.OST_DataDevices).ToElements()


def _get_all_light_devices(doc):
    return _collect_by_category(doc, DB.BuiltInCategory.OST_LightingDevices).ToElements()


def _get_all_light_fixtures(doc):
    return _collect_by_category(doc, DB.BuiltInCategory.OST_LightingFixtures).ToElements()


def _get_all_mech_control_devices(doc):
    return _collect_by_category(doc, DB.BuiltInCategory.OST_MechanicalControlDevices).ToElements()


def _load_module_from_path(module_name, file_path):
    try:
        import importlib.util

        spec = importlib.util.spec_from_file_location(module_name, file_path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    except Exception:
        import imp

        return imp.load_source(module_name, file_path)


def _get_supercircuit_v5_dir():
    return os.path.normpath(
        os.path.join(os.path.dirname(__file__), "..", "SuperCircuitV5.pushbutton")
    )


def _load_supercircuit_circuits():
    v5_dir = _get_supercircuit_v5_dir()
    if not os.path.isdir(v5_dir):
        raise IOError("SuperCircuitV5 folder not found: {}".format(v5_dir))

    if v5_dir not in sys.path:
        sys.path.insert(0, v5_dir)

    circuits_path = os.path.join(v5_dir, "circuits.py")
    if not os.path.isfile(circuits_path):
        raise IOError("SuperCircuitV5 circuits.py not found: {}".format(circuits_path))

    return _load_module_from_path("supercircuitv5_circuits_audit", circuits_path)


def _select_client():
    selection = forms.CommandSwitchWindow.show(
        list(CLIENT_CHOICES.keys()),
        message="Select client configuration",
    )
    if not selection:
        return None
    return CLIENT_CHOICES.get(selection)


def _load_client_helpers(client_key):
    if client_key == "planet_fitness":
        try:
            from PFlib import PFhelpers
            return PFhelpers
        except ImportError as ex:
            logger.warning("PF helpers unavailable: {}".format(ex))
    elif client_key == "heb":
        try:
            from HEBlib import HEBhelper
            try:
                import importlib

                importlib.reload(HEBhelper)
            except Exception:
                try:
                    reload(HEBhelper)  # IronPython fallback
                except Exception:
                    pass
            return HEBhelper
        except ImportError as ex:
            logger.warning("HEB helpers unavailable: {}".format(ex))
    return None


def _select_scope():
    current_selection = list(revit.get_selection() or [])
    if not current_selection:
        return False

    choice = forms.CommandSwitchWindow.show(
        list(SCOPE_CHOICES.keys()),
        message="Select audit scope",
    )
    if not choice:
        return None
    return SCOPE_CHOICES.get(choice, False)


def _collect_elements(circuits, doc, selection_only=False):
    collectors = (
        _get_all_elec_fixtures,
        _get_all_data_devices,
        _get_all_mech_control_devices,
    )
    selection_getter = revit.get_selection if selection_only else (lambda: [])
    return circuits.collect_target_elements(doc, collectors, selection_getter, logger)


def _filter_disallowed_elements(elements):
    filtered = []
    skipped = 0
    for element in elements or []:
        category = getattr(element, "Category", None)
        category_id = category.Id.IntegerValue if category and category.Id else None
        if category_id in EXCLUDED_CATEGORY_IDS:
            skipped += 1
            continue
        filtered.append(element)

    if skipped:
        logger.info("Skipped {} lighting element(s); use the dedicated lighting tool.".format(skipped))
    return filtered


def _run_client_preprocess(info_items, doc, panel_lookup, client_helpers):
    if client_helpers and hasattr(client_helpers, "preprocess_items"):
        try:
            processed = client_helpers.preprocess_items(info_items, doc, panel_lookup, logger)
            if processed:
                return list(processed)
        except Exception as ex:
            logger.warning("Client preprocess_items failed: {}".format(ex))
    return info_items


def _group_priority(group_type):
    priority_map = {
        "dedicated": 0,
        "position": 1,
        "special": 2,
    }
    return priority_map.get(group_type or "normal", 3)


def _load_priority(group, group_priority_value, client_helpers):
    if group_priority_value < 3:
        return ""

    module = client_helpers
    if module and hasattr(module, "get_load_priority"):
        try:
            return module.get_load_priority(group)
        except Exception:
            return ""
    return ""


def _sort_groups(circuits, groups, client_helpers):
    def sort_key(group):
        priority = _group_priority(group.get("group_type"))
        panel = (group.get("panel_name") or "").lower()
        load_priority = _load_priority(group, priority, client_helpers)
        circuit_number = group.get("circuit_number")
        circuit_sort = circuits.try_parse_int(circuit_number)
        if circuit_sort is None:
            circuit_sort = circuit_number or group.get("key") or ""
        return (
            priority,
            panel,
            load_priority,
            circuit_sort,
            group.get("key"),
        )

    return sorted(groups, key=sort_key)


def _normalize_text(value):
    if value is None:
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    return text.upper()


def _get_param_string(element, bip=None, lookup_name=None):
    param = None
    if bip is not None:
        try:
            param = element.get_Parameter(bip)
        except Exception:
            param = None
    if (not param) and lookup_name:
        try:
            param = element.LookupParameter(lookup_name)
        except Exception:
            param = None
    if not param:
        return ""

    try:
        val = query.get_param_value(param)
    except Exception:
        val = None
    if val is None:
        try:
            val = param.AsString()
        except Exception:
            val = None
    return (str(val).strip() if val is not None else "")


def _iter_power_systems(circuits, element):
    if not element:
        return []

    mep_model = getattr(element, "MEPModel", None)
    if not mep_model:
        return []

    systems = getattr(mep_model, "ElectricalSystems", None)
    if not systems:
        return []

    power_type = DB.Electrical.ElectricalSystemType.PowerCircuit
    power_systems = []
    for system in circuits.iterate_collection(systems):
        if not system:
            continue
        include = False
        for attr in ("SystemType", "ElectricalSystemType"):
            try:
                if getattr(system, attr, None) == power_type:
                    include = True
                    break
            except Exception:
                continue
        if include:
            power_systems.append(system)
    return power_systems


def _is_power_circuit_system(system):
    if not system:
        return False
    power_type = DB.Electrical.ElectricalSystemType.PowerCircuit
    for attr in ("SystemType", "ElectricalSystemType"):
        try:
            if getattr(system, attr, None) == power_type:
                return True
        except Exception:
            continue
    return False


def _collect_system_member_ids(circuits, system):
    member_ids = set()
    if not system:
        return member_ids

    elements = getattr(system, "Elements", None)
    for element in circuits.iterate_collection(elements):
        elem_id = getattr(getattr(element, "Id", None), "IntegerValue", None)
        if elem_id is not None:
            member_ids.add(elem_id)
    return member_ids


def _build_power_circuit_index(doc, circuits):
    by_element_id = {}
    system_description_cache = {}
    system_count = 0

    collector = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_ElectricalCircuit)
        .WhereElementIsNotElementType()
    )

    for system in collector:
        if not _is_power_circuit_system(system):
            continue
        system_id = getattr(getattr(system, "Id", None), "IntegerValue", None)
        desc = system_description_cache.get(system_id) if system_id is not None else None
        if desc is None:
            desc = _describe_system(system)
            if system_id is not None:
                system_description_cache[system_id] = desc
        if not desc:
            continue

        member_ids = _collect_system_member_ids(circuits, system)
        if not member_ids:
            continue

        system_count += 1
        for elem_id in member_ids:
            by_element_id.setdefault(elem_id, []).append(desc)

    return by_element_id, system_count


def _describe_system(system):
    if not system:
        return None

    circuit_number = ""
    try:
        circuit_number = str(getattr(system, "CircuitNumber", "") or "").strip()
    except Exception:
        circuit_number = ""
    if not circuit_number:
        bip = getattr(DB.BuiltInParameter, "RBS_ELEC_CIRCUIT_NUMBER", None)
        circuit_number = _get_param_string(system, bip=bip)

    load_name = _get_param_string(
        system,
        bip=DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME,
    )

    panel_name = _get_param_string(
        system,
        bip=DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
    )
    if not panel_name:
        try:
            panel_name = str(getattr(system, "PanelName", "") or "").strip()
        except Exception:
            panel_name = ""
    if not panel_name:
        base_equipment = getattr(system, "BaseEquipment", None)
        if base_equipment:
            panel_name = _get_param_string(base_equipment, lookup_name="Panel Name")

    sys_id = getattr(getattr(system, "Id", None), "IntegerValue", None)

    return {
        "system": system,
        "system_id": sys_id,
        "panel_name": panel_name or "",
        "circuit_number": circuit_number or "",
        "load_name": load_name or "",
    }


def _match_expected_to_actual(expected, actual_descriptions):
    exp_panel = _normalize_text(expected.get("panel_name"))
    exp_load = _normalize_text(expected.get("load_name"))

    panel_load_matches = []
    panel_matches = []
    load_matches = []
    for actual in actual_descriptions:
        act_panel = _normalize_text(actual.get("panel_name"))
        act_load = _normalize_text(actual.get("load_name"))
        if act_panel == exp_panel:
            panel_matches.append(actual)
        if act_load == exp_load:
            load_matches.append(actual)
        if act_panel == exp_panel and act_load == exp_load:
            panel_load_matches.append(actual)

    if len(actual_descriptions) == 0:
        return False, None, "No power circuit exists in model"
    if panel_load_matches and len(actual_descriptions) == 1:
        return True, panel_load_matches[0], ""
    if panel_load_matches and len(actual_descriptions) > 1:
        return False, panel_load_matches[0], "Element is assigned to multiple power circuits"
    if panel_matches:
        return False, panel_matches[0], "Load name differs from expected"
    if load_matches:
        return False, load_matches[0], "Panel differs from expected"
    return False, actual_descriptions[0], "Panel and load name differ from expected"


def _format_circuit_ref(panel_name, circuit_number, load_name):
    panel = panel_name or "<blank panel>"
    circuit = circuit_number or "<blank circuit>"
    load = load_name or "<blank load>"
    return "{} / {} / {}".format(panel, circuit, load)


def _format_panel_load_ref(panel_name, load_name):
    panel = panel_name or "<blank panel>"
    load = load_name or "<blank load>"
    return "{} / {}".format(panel, load)


def _format_panel_load_circuit_ref(panel_name, load_name, circuit_number):
    panel = panel_name or "<blank panel>"
    load = load_name or "<blank load>"
    circuit = circuit_number or "<blank circuit>"
    return "{} / {} / {}".format(panel, load, circuit)


def _linkify_element(element_id):
    if element_id is None:
        return "N/A"
    try:
        return output.linkify(DB.ElementId(int(element_id)))
    except Exception:
        return str(element_id)


def _safe_text(value):
    if value is None:
        return ""
    try:
        return str(value).strip()
    except Exception:
        return ""


_DETAIL_GRID_COLUMNS = [
    ("host_id", "Host Element"),
    ("circuit_elem_id", "Circuit Element"),
    ("expected_ref", "SuperCircuitV5 Panel/Load"),
    ("ckt_panel_cedt_raw", "CKT_Panel_CEDT Raw"),
    ("ckt_circuit_raw", "CKT_Circuit Number_CEDT Raw"),
    ("actual_ref", "Actual Panel/Load"),
    ("reason", "Reason"),
]


_FILTER_TEXTBOX_NAMES = [
    "FilterHostElement",
    "FilterCircuitElement",
    "FilterExpected",
    "FilterPanelCedt",
    "FilterCircuitCedt",
    "FilterActual",
    "FilterReason",
]


def _rowfilter_column_name(column_name):
    return "[{}]".format((column_name or "").replace("]", "]]"))


def _rowfilter_escape_value(value):
    text = _safe_text(value)
    text = text.replace("'", "''")
    text = text.replace("[", "[[]")
    text = text.replace("]", "[]]")
    text = text.replace("%", "[%]")
    text = text.replace("*", "[*]")
    return text


class ChangedElementDetailsWindow(forms.WPFWindow):
    def __init__(self, xaml_path, rows, total_count=None):
        forms.WPFWindow.__init__(self, xaml_path)
        self.rows = list(rows or [])
        self.total_count = int(total_count if total_count is not None else len(self.rows))
        self._filters = {key: "" for key, _header in _DETAIL_GRID_COLUMNS}
        self._sort_column = None
        self._sort_direction = "ASC"
        self._data_table = self._build_data_table(self.rows)
        self._data_view = self._data_table.DefaultView

        grid = self.FindName("DetailsGrid")
        if grid is not None:
            grid.ItemsSource = self._data_view

        self._update_summary()

    def _build_data_table(self, rows):
        if DataTable is None:
            raise Exception("System.Data is unavailable in this environment.")

        table = DataTable("ChangedElementDetails")
        for key, _header in _DETAIL_GRID_COLUMNS:
            table.Columns.Add(key, str)

        for row in rows:
            data_row = table.NewRow()
            for key, _header in _DETAIL_GRID_COLUMNS:
                data_row[key] = _safe_text((row or {}).get(key))
            table.Rows.Add(data_row)
        return table

    def _update_summary(self):
        summary = self.FindName("SummaryText")
        if summary is None:
            return
        shown = int(getattr(self._data_view, "Count", 0))
        summary.Text = "Showing {} of {} changed rows.".format(shown, self.total_count)

    def _apply_view(self):
        filters = []
        for key, filter_text in self._filters.items():
            text = _safe_text(filter_text)
            if not text:
                continue
            expr = "CONVERT({}, 'System.String') LIKE '%{}%'".format(
                _rowfilter_column_name(key),
                _rowfilter_escape_value(text),
            )
            filters.append(expr)

        self._data_view.RowFilter = " AND ".join(filters) if filters else ""
        if self._sort_column:
            self._data_view.Sort = "{} {}".format(
                _rowfilter_column_name(self._sort_column),
                self._sort_direction,
            )
        else:
            self._data_view.Sort = ""
        self._update_summary()

    def OnFilterChanged(self, sender, args):
        key = _safe_text(getattr(sender, "Tag", None))
        if key in self._filters:
            self._filters[key] = _safe_text(getattr(sender, "Text", ""))
            self._apply_view()

    def OnGridSorting(self, sender, args):
        column_name = _safe_text(getattr(args.Column, "SortMemberPath", None)) or _safe_text(
            getattr(args.Column, "Header", None)
        )
        if not column_name:
            return

        if self._sort_column == column_name:
            self._sort_direction = "DESC" if self._sort_direction == "ASC" else "ASC"
        else:
            self._sort_column = column_name
            self._sort_direction = "ASC"

        self._apply_view()

        grid = self.FindName("DetailsGrid")
        if grid is not None:
            for column in grid.Columns:
                column.SortDirection = None
        if ListSortDirection is not None:
            args.Column.SortDirection = (
                ListSortDirection.Ascending
                if self._sort_direction == "ASC"
                else ListSortDirection.Descending
            )
        args.Handled = True

    def OnClearFilters(self, sender, args):
        for column_name in self._filters.keys():
            self._filters[column_name] = ""

        for textbox_name in _FILTER_TEXTBOX_NAMES:
            textbox = self.FindName(textbox_name)
            if textbox is not None:
                textbox.Text = ""

        self._sort_column = None
        self._sort_direction = "ASC"
        grid = self.FindName("DetailsGrid")
        if grid is not None:
            for column in grid.Columns:
                column.SortDirection = None

        self._apply_view()

    def OnCloseWindow(self, sender, args):
        self.Close()


def _show_changed_details_window(rows, total_count=None):
    if DataTable is None:
        forms.alert(
            "Changed details grid requires System.Data, which is unavailable in this runtime.",
            title=__title__,
        )
        return

    xaml_path = script.get_bundle_file("ChangedElementDetailsWindow.xaml")
    if not xaml_path:
        xaml_path = os.path.join(os.path.dirname(__file__), "ChangedElementDetailsWindow.xaml")
    if not os.path.exists(xaml_path):
        logger.warning("Changed details XAML not found: %s", xaml_path)
        return

    try:
        window = ChangedElementDetailsWindow(xaml_path, rows, total_count=total_count)
        window.show_dialog()
    except Exception as ex:
        logger.warning("Failed to open changed details window: %s", ex)


def _build_audit(circuits, groups):
    group_results = []
    changed_member_rows = []
    doc = revit.doc
    circuit_index, indexed_system_count = _build_power_circuit_index(doc, circuits)
    logger.info(
        "Audit Circuits: indexed {} power circuits across {} connected elements.".format(
            indexed_system_count, len(circuit_index)
        )
    )

    actual_by_member_key = {}
    system_description_cache = {}

    for group in groups or []:
        members = list(group.get("members") or [])
        if not members:
            continue

        expected_panel = group.get("panel_name") or ""
        expected_circuit = group.get("circuit_number") or ""
        expected_load = group.get("load_name") or ""

        matched_count = 0
        actual_system_ids = set()

        for member in members:
            host_element = member.get("element")
            circuit_element = member.get("circuit_element") or host_element
            host_id = getattr(getattr(host_element, "Id", None), "IntegerValue", None)
            circuit_elem_id = getattr(getattr(circuit_element, "Id", None), "IntegerValue", None)

            expected = {
                "panel_name": member.get("panel_name") or expected_panel,
                "circuit_number": member.get("circuit_number") or expected_circuit,
                "load_name": member.get("load_name") or expected_load,
            }

            cache_key = (host_id, circuit_elem_id)
            cached = actual_by_member_key.get(cache_key)
            if cached is None:
                actual_descriptions = []
                seen_system_ids = set()
                cached_system_ids = []

                candidate_ids = []
                if circuit_elem_id is not None:
                    candidate_ids.append(circuit_elem_id)
                if host_id is not None and host_id not in candidate_ids:
                    candidate_ids.append(host_id)

                for candidate_id in candidate_ids:
                    for desc in circuit_index.get(candidate_id, []):
                        system_id = desc.get("system_id")
                        if system_id in seen_system_ids:
                            continue
                        seen_system_ids.add(system_id)
                        actual_descriptions.append(desc)
                        if system_id is not None:
                            cached_system_ids.append(system_id)

                if not actual_descriptions:
                    power_systems = _iter_power_systems(circuits, circuit_element)
                    for system in power_systems:
                        system_id = getattr(getattr(system, "Id", None), "IntegerValue", None)
                        desc = system_description_cache.get(system_id) if system_id is not None else None
                        if desc is None:
                            desc = _describe_system(system)
                            if system_id is not None:
                                system_description_cache[system_id] = desc
                        if not desc:
                            continue
                        if system_id in seen_system_ids:
                            continue
                        seen_system_ids.add(system_id)
                        actual_descriptions.append(desc)
                        if desc.get("system_id") is not None:
                            cached_system_ids.append(desc.get("system_id"))

                cached = (actual_descriptions, cached_system_ids)
                actual_by_member_key[cache_key] = cached

            actual_descriptions, cached_system_ids = cached
            for sys_id in cached_system_ids:
                actual_system_ids.add(sys_id)

            is_match, matched_actual, reason = _match_expected_to_actual(expected, actual_descriptions)
            if is_match:
                matched_count += 1

            if not is_match:
                actual_ref = "None"
                if matched_actual:
                    actual_ref = _format_panel_load_circuit_ref(
                        matched_actual.get("panel_name"),
                        matched_actual.get("load_name"),
                        matched_actual.get("circuit_number"),
                    )
                elif actual_descriptions:
                    first = actual_descriptions[0]
                    actual_ref = _format_panel_load_circuit_ref(
                        first.get("panel_name"),
                        first.get("load_name"),
                        first.get("circuit_number"),
                    )

                changed_member_rows.append(
                    {
                        "host_id": host_id,
                        "circuit_elem_id": circuit_elem_id,
                        "expected_ref": _format_panel_load_ref(
                            expected.get("panel_name"),
                            expected.get("load_name"),
                        ),
                        "actual_ref": actual_ref,
                        "ckt_panel_cedt_raw": _get_param_string(
                            host_element, lookup_name="CKT_Panel_CEDT"
                        ) or "",
                        "ckt_circuit_raw": _get_param_string(
                            host_element, lookup_name="CKT_Circuit Number_CEDT"
                        ) or "",
                        "reason": reason,
                    }
                )

        is_group_match = matched_count == len(members)
        group_results.append(
            {
                "group_key": group.get("key") or "",
                "group_type": group.get("group_type") or "normal",
                "expected_panel": expected_panel,
                "expected_circuit": expected_circuit,
                "expected_load": expected_load,
                "member_count": len(members),
                "matched_members": matched_count,
                "is_match": is_group_match,
                "actual_system_ids": sorted(actual_system_ids),
            }
        )

    return group_results, changed_member_rows


def _render_report(
    group_results,
    changed_member_rows,
    include_changed_details=True,
    detail_row_cap=500,
):
    output.close_others()
    output.print_md("# Audit Circuits")

    total_groups = len(group_results)
    total_members = sum(r.get("member_count", 0) for r in group_results)
    matched_groups = [r for r in group_results if r.get("is_match")]
    changed_groups = [r for r in group_results if not r.get("is_match")]
    matched_members = sum(r.get("matched_members", 0) for r in group_results)
    changed_members = total_members - matched_members

    output.print_md("## Summary")
    output.print_md("- Expected circuit groups (SuperCircuitV5 logic): **{}**".format(total_groups))
    output.print_md("- Matched groups: **{}**".format(len(matched_groups)))
    output.print_md("- Changed groups: **{}**".format(len(changed_groups)))
    output.print_md("- Expected elements audited: **{}**".format(total_members))
    output.print_md("- Matched elements: **{}**".format(matched_members))
    output.print_md("- Changed elements: **{}**".format(changed_members))

    matched_rows = []
    changed_rows = []
    for result in group_results:
        expected_ref = _format_panel_load_ref(
            result.get("expected_panel"),
            result.get("expected_load"),
        )
        actual_ids = result.get("actual_system_ids") or []
        actual_links = ", ".join([_linkify_element(sys_id) for sys_id in actual_ids]) if actual_ids else "None"
        row = [
            result.get("group_key") or "",
            expected_ref,
            str(result.get("member_count", 0)),
            str(result.get("matched_members", 0)),
            actual_links,
        ]
        if result.get("is_match"):
            matched_rows.append(row)
        else:
            changed_rows.append(row)

    output.print_md("## Matched Circuits")
    if matched_rows:
        output.print_table(
            table_data=matched_rows,
            columns=["Group Key", "Expected Panel/Load", "Members", "Matched", "Actual Circuit Id(s)"],
        )
    else:
        output.print_md("_No fully matched circuit groups found._")

    output.print_md("## Changed Circuits")
    if changed_rows:
        output.print_table(
            table_data=changed_rows,
            columns=["Group Key", "Expected Panel/Load", "Members", "Matched", "Actual Circuit Id(s)"],
        )
    else:
        output.print_md("_No changed circuit groups found._")

    output.print_md("## Changed Element Details")
    if changed_member_rows and include_changed_details:
        max_rows = max(1, int(detail_row_cap or 120))
        shown = min(len(changed_member_rows), max_rows)
        output.print_md(
            "Opening sortable/filterable WPF grid window for changed details (**{}** row(s)).".format(shown)
        )
        if len(changed_member_rows) > max_rows:
            output.print_md(
                "_Showing first {} changed rows out of {}._".format(max_rows, len(changed_member_rows))
            )
    elif changed_member_rows and (not include_changed_details):
        output.print_md(
            "_Detailed rows skipped for speed ({} changed elements). Re-run and choose detailed mode if needed._".format(
                len(changed_member_rows)
            )
        )
    else:
        output.print_md("_No element-level mismatches detected._")


def main():
    started_at = time.time()
    doc = revit.doc
    if not doc:
        forms.alert("No active Revit document found.", title=__title__)
        return

    try:
        circuits = _load_supercircuit_circuits()
    except Exception as ex:
        forms.alert("Failed to load SuperCircuitV5 logic:\n{}".format(ex), title=__title__)
        return

    client_key = _select_client()
    if not client_key:
        return
    client_helpers = _load_client_helpers(client_key)

    selection_only = _select_scope()
    if selection_only is None:
        return

    t0 = time.time()
    panels = list(_get_all_panels(doc))
    panel_lookup = circuits.build_panel_lookup(panels)
    logger.info("Audit Circuits timing: panel lookup {:.2f}s".format(time.time() - t0))

    t1 = time.time()
    elements = _collect_elements(circuits, doc, selection_only=selection_only)
    elements = _filter_disallowed_elements(elements)
    if not elements:
        forms.alert("No eligible elements found for audit.", title=__title__)
        return
    logger.info(
        "Audit Circuits timing: collected {} candidate elements in {:.2f}s".format(
            len(elements), time.time() - t1
        )
    )

    t2 = time.time()
    info_items = circuits.gather_element_info(doc, elements, panel_lookup, logger)
    if not info_items:
        forms.alert("No elements with CKT_Panel_CEDT/CKT_Circuit Number_CEDT data were found.", title=__title__)
        return
    logger.info(
        "Audit Circuits timing: gathered {} info items in {:.2f}s".format(
            len(info_items), time.time() - t2
        )
    )

    t3 = time.time()
    info_items = _run_client_preprocess(info_items, doc, panel_lookup, client_helpers)
    groups = circuits.assemble_groups(info_items, client_helpers, logger)
    if not groups:
        forms.alert("SuperCircuitV5 grouping produced no circuit groups.", title=__title__)
        return
    logger.info(
        "Audit Circuits timing: preprocess+grouping produced {} groups in {:.2f}s".format(
            len(groups), time.time() - t3
        )
    )

    t4 = time.time()
    groups = _sort_groups(circuits, groups, client_helpers)
    group_results, changed_member_rows = _build_audit(circuits, groups)
    logger.info(
        "Audit Circuits timing: audit comparison completed in {:.2f}s".format(time.time() - t4)
    )

    include_details = True
    detail_cap = 500
    changed_count = len(changed_member_rows)
    if changed_count > 120:
        mode = forms.alert(
            "{} changed elements found.\n\nDetailed printing can be slow.\nChoose report mode:".format(changed_count),
            title=__title__,
            options=["Fast Summary", "Include Details"],
        )
        include_details = (mode == "Include Details")
        if not include_details:
            detail_cap = 0

    _render_report(
        group_results,
        changed_member_rows,
        include_changed_details=include_details,
        detail_row_cap=detail_cap,
    )

    if include_details and changed_member_rows:
        max_rows = max(1, int(detail_cap or 500))
        _show_changed_details_window(
            changed_member_rows[:max_rows],
            total_count=len(changed_member_rows),
        )

    logger.info("Audit Circuits timing: total {:.2f}s".format(time.time() - started_at))

    forms.alert(
        "Audit complete.\n\nSee output panel for matched vs changed circuits.",
        title=__title__,
    )


if __name__ == "__main__":
    main()
