# -*- coding: utf-8 -*-
__title__ = "SUPER CIRCUIT V5"

from collections import OrderedDict
import logging
import os
import re

from pyrevit import revit, script, forms, DB
from System.Windows.Media import Color, SolidColorBrush

from Snippets._elecutils import (
    get_all_data_devices,
    get_all_elec_fixtures,
    get_all_light_devices,
    get_all_light_fixtures,
    get_all_mech_control_devices,
    get_all_panels,
)

import circuits

logger = script.get_logger()
logger.setLevel(logging.INFO)

try:
    basestring
except NameError:  # Python 3 fallback
    basestring = str

CLIENT_CHOICES = OrderedDict([
    ("Planet Fitness", "planet_fitness"),
    ("HEB", "heb"),
])

SCOPE_CHOICES = OrderedDict([
    ("All eligible elements", False),
    ("Selected elements only", True),
])

client_helpers = None
EXCLUDED_CATEGORY_IDS = {
    DB.ElementId(DB.BuiltInCategory.OST_LightingDevices).IntegerValue,
    DB.ElementId(DB.BuiltInCategory.OST_LightingFixtures).IntegerValue,
}


DEFAULT_ROW_FOREGROUND = SolidColorBrush(Color.FromRgb(0, 0, 0))
COLOR_GREEN = SolidColorBrush(Color.FromRgb(46, 125, 50))
COLOR_BLUE = SolidColorBrush(Color.FromRgb(21, 101, 192))
COLOR_INDIGO = SolidColorBrush(Color.FromRgb(48, 63, 159))
COLOR_RED = SolidColorBrush(Color.FromRgb(198, 40, 40))
COLOR_LIGHT_TEXT = SolidColorBrush(Color.FromRgb(255, 255, 255))


def _group_row_colors(group_size):
    try:
        size = int(group_size or 0)
    except Exception:
        size = 0

    if size == 2:
        return COLOR_GREEN, COLOR_LIGHT_TEXT
    if size == 3:
        return COLOR_BLUE, COLOR_LIGHT_TEXT
    if 4 <= size <= 7:
        return COLOR_INDIGO, COLOR_LIGHT_TEXT
    if size >= 8:
        return COLOR_RED, COLOR_LIGHT_TEXT
    return None, DEFAULT_ROW_FOREGROUND

class PreviewRow(object):
    def __init__(
        self,
        group_index,
        group_label,
        panel_name,
        circuit_number,
        family_type,
        element_id,
        source_item,
        row_background=None,
        row_foreground=None,
        is_spacer=False,
        panel_options=None,
        circuit_options=None,
    ):
        self.group_index = group_index
        self.group_label = group_label
        self.panel_name = panel_name or ""
        self.circuit_number = circuit_number or ""
        self.family_type = family_type or ""
        self.element_id = str(element_id) if element_id is not None else ""
        self.source_item = source_item
        self.row_background = row_background
        self.row_foreground = row_foreground or DEFAULT_ROW_FOREGROUND
        self.is_spacer = bool(is_spacer)
        self.panel_options = list(panel_options or [])
        self.circuit_options = list(circuit_options or [])

class SuperCircuitPreviewWindow(forms.WPFWindow):
    def __init__(self, xaml_path, rows):
        forms.WPFWindow.__init__(self, xaml_path)
        self.rows = rows or []
        self.accepted = False

        header = self.FindName("HeaderText")
        if header is not None:
            header.Text = (
                "Preview of circuits to be created. Rows are grouped by circuit batch and panel. "
                "Edit panel and circuit directly in the combo boxes in this window, then click Run Circuits. "
                "Use the Color Key at the top (1=No Color, 2=Green, 3=Blue, 4-7=Indigo, 8+=Red). Blank rows separate each circuit batch."
            )

        summary = self.FindName("SummaryText")
        if summary is not None:
            data_rows = [row for row in self.rows if not getattr(row, "is_spacer", False)]
            summary.Text = "{} element(s) in {} circuit batch(es).".format(
                len(data_rows),
                len({row.group_index for row in data_rows}),
            )

        grid = self.FindName("PreviewGrid")
        if grid is not None:
            grid.ItemsSource = self.rows

        edit_btn = self.FindName("EditSelectedButton")
        run_btn = self.FindName("RunButton")
        cancel_btn = self.FindName("CancelButton")

        if edit_btn is not None:
            edit_btn.Click += self._on_edit_selected
        if run_btn is not None:
            run_btn.Click += self._on_run
        if cancel_btn is not None:
            cancel_btn.Click += self._on_cancel

    def _selected_rows(self):
        grid = self.FindName("PreviewGrid")
        if grid is None:
            return []
        selected = []
        try:
            for item in grid.SelectedItems:
                if getattr(item, "is_spacer", False):
                    continue
                if not getattr(item, "source_item", None):
                    continue
                selected.append(item)
        except Exception:
            pass
        return selected

    def _apply_row_values(self, rows, panel_value, circuit_value):
        for row in rows:
            if getattr(row, "is_spacer", False):
                continue
            row.panel_name = panel_value
            row.circuit_number = circuit_value
        grid = self.FindName("PreviewGrid")
        if grid is not None:
            try:
                grid.Items.Refresh()
            except Exception:
                pass

    def _on_edit_selected(self, sender, args):
        grid = self.FindName("PreviewGrid")
        if grid is None:
            return
        try:
            grid.Focus()
            grid.BeginEdit()
        except Exception:
            pass

    def _close_with_result(self, accepted):
        self.accepted = bool(accepted)
        try:
            self.DialogResult = bool(accepted)
        except Exception:
            pass
        self.Close()

    def _on_run(self, sender, args):
        self._close_with_result(True)

    def _on_cancel(self, sender, args):
        self._close_with_result(False)


def _select_client():
    selection = forms.CommandSwitchWindow.show(
        list(CLIENT_CHOICES.keys()), message="Select client configuration"
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
                logger.info("Reloaded HEBhelper module.")
            except Exception:
                try:
                    reload(HEBhelper)  # IronPython fallback
                    logger.info("Reloaded HEBhelper module via reload().")
                except Exception as ex:
                    logger.warning("HEBhelper reload failed: {}".format(ex))

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
        message="Select circuiting scope",
    )
    if not choice:
        return None
    return SCOPE_CHOICES.get(choice, False)


def _split_panel_choices(value):
    if not value:
        return []
    text = value if isinstance(value, basestring) else str(value)
    for sep in (",", ";", "|", "/", "\n", "\r"):
        text = text.replace(sep, " ")
    candidates = [part.strip() for part in text.split(" ") if part.strip()]
    unique = []
    seen = set()
    for name in candidates:
        upper = name.upper()
        if upper in seen:
            continue
        seen.add(upper)
        unique.append(name)
    return unique


def _tokenize_panel_value(value):
    if not value:
        return []
    text = value if isinstance(value, basestring) else str(value)
    tokens = re.findall(r"[A-Za-z0-9]+", text)
    return [token.strip() for token in tokens if token.strip()]


def _get_point(elem):
    if not elem:
        return None
    location = getattr(elem, "Location", None)
    if location:
        point = getattr(location, "Point", None)
        if point:
            return point
        curve = getattr(location, "Curve", None)
        if curve:
            try:
                return curve.Evaluate(0.5, True)
            except Exception:
                pass
    try:
        bbox = elem.get_BoundingBox(None)
    except Exception:
        bbox = None
    if bbox:
        return DB.XYZ(
            (bbox.Min.X + bbox.Max.X) * 0.5,
            (bbox.Min.Y + bbox.Max.Y) * 0.5,
            (bbox.Min.Z + bbox.Max.Z) * 0.5,
        )
    return None


def _distance(point_a, point_b):
    if point_a is None or point_b is None:
        return None
    try:
        return point_a.DistanceTo(point_b)
    except Exception:
        try:
            dx = point_a.X - point_b.X
            dy = point_a.Y - point_b.Y
            dz = point_a.Z - point_b.Z
            return (dx * dx + dy * dy + dz * dz) ** 0.5
        except Exception:
            return None


def _build_panel_point_cache(panel_lookup):
    cache = {}
    for name, info in (panel_lookup or {}).items():
        panel_elem = info.get("element")
        point = _get_point(panel_elem)
        if not point:
            continue
        upper = (name or "").strip().upper()
        if not upper or upper in cache:
            continue
        cache[upper] = {"name": name, "point": point}
    return cache


def _emit_debug(output, text):
    try:
        output.print_md(text)
    except Exception:
        pass
    try:
        print(text)
    except Exception:
        pass


def _debug_ba_da(label, items, panel_lookup, output):
    panel_point_cache = _build_panel_point_cache(panel_lookup)
    for item in items:
        panel_value = item.get("panel_name")
        tokens = _tokenize_panel_value(panel_value)
        token_set = {token.upper() for token in tokens}
        if not (("BA" in token_set) or ("DA" in token_set)):
            continue
        element = item.get("element")
        elem_id = getattr(getattr(element, "Id", None), "IntegerValue", None)
        location = item.get("location")
        load_name = item.get("load_name") or "None"
        if location:
            loc_text = "{:.3f},{:.3f},{:.3f}".format(location.X, location.Y, location.Z)
        else:
            loc_text = "None"
        _emit_debug(
            output,
            "INFO SuperCircuitV5 {} BA/DA debug | element {} | location {} | CKT_Panel_CEDT {} | CKT_Load Name_CEDT {} | tokens {}".format(
                label,
                elem_id if elem_id is not None else "unknown",
                loc_text,
                panel_value or "None",
                load_name,
                ", ".join(sorted(token_set)) if token_set else "none",
            ),
        )
        for panel_key in ("BA", "DA"):
            if panel_key not in token_set:
                _emit_debug(
                    output,
                    "INFO SuperCircuitV5 {} BA/DA debug | panel {} not listed in CKT_Panel_CEDT".format(
                        label, panel_key
                    ),
                )
                continue
            entry = panel_point_cache.get(panel_key)
            if not entry:
                _emit_debug(
                    output,
                    "INFO SuperCircuitV5 {} BA/DA debug | panel {} missing from cache".format(
                        label, panel_key
                    ),
                )
                continue
            point = entry.get("point")
            dist = _distance(location, point) if location and point else None
            if point:
                ptext = "{:.3f},{:.3f},{:.3f}".format(point.X, point.Y, point.Z)
            else:
                ptext = "None"
            _emit_debug(
                output,
                "INFO SuperCircuitV5 {} BA/DA debug | panel {} point {} | distance {}".format(
                    label,
                    panel_key,
                    ptext,
                    "{:.3f}".format(dist) if dist is not None else "None",
                ),
            )


def _collect_elements(doc, selection_only=False):
    collectors = (
        get_all_elec_fixtures,
        get_all_light_devices,
        get_all_light_fixtures,
        get_all_data_devices,
        get_all_mech_control_devices,
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


def _run_creation(doc, groups):
    return circuits.run_creation(
        doc,
        groups,
        lambda d, g: circuits.create_circuit(d, g, logger),
        logger,
        transaction_label="SuperCircuitV5 - Create Circuits",
    )


def _run_apply_data(doc, created_systems):
    circuits.run_apply_data(
        doc,
        created_systems,
        lambda system, group: circuits.apply_circuit_data(system, group, logger),
        logger,
        transaction_label="SuperCircuitV5 - Apply Circuit Data",
    )


def _group_priority(group_type):
    priority_map = {
        "dedicated": 0,
        "nongrouped": 1,
        "special": 2,
        "position": 2,
    }
    return priority_map.get(group_type or "normal", 3)


def _load_priority(group, group_priority_value):
    if group_priority_value < 3:
        return 0
    if client_helpers and hasattr(client_helpers, "get_load_priority"):
        try:
            return client_helpers.get_load_priority(group)
        except Exception as ex:
            logger.warning("Client load priority lookup failed: {}".format(ex))
    return 99


def _sort_groups(groups):
    def sort_key(group):
        members = group.get("members") or []
        choice_counts = [m.get("panel_choice_count") for m in members if m.get("panel_choice_count") is not None]
        panel_choice_count = min(choice_counts) if choice_counts else 99
        panel_choice_priority = 0 if panel_choice_count == 1 else 1

        priority = _group_priority(group.get("group_type"))
        panel = (group.get("panel_name") or "").lower()
        load_priority = _load_priority(group, priority)
        circuit_number = group.get("circuit_number")
        circuit_sort = circuits.try_parse_int(circuit_number)
        if circuit_sort is None:
            circuit_sort = circuit_number or group.get("key") or ""
        return (
            panel_choice_priority,
            panel_choice_count,
            priority,
            panel,
            load_priority,
            circuit_sort,
            group.get("key"),
        )

    return sorted(groups, key=sort_key)


def _family_type_text(element):
    if not element:
        return "Unknown : Unknown"

    family_name = None
    type_name = None

    try:
        symbol = getattr(element, "Symbol", None)
        if symbol:
            family = getattr(symbol, "Family", None)
            family_name = getattr(family, "Name", None)
            type_name = getattr(symbol, "Name", None)
    except Exception:
        pass

    if not family_name:
        try:
            fam_param = element.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
            if fam_param and fam_param.HasValue:
                family_name = fam_param.AsValueString() or fam_param.AsString()
        except Exception:
            pass

    if not type_name:
        try:
            type_param = element.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM)
            if type_param and type_param.HasValue:
                type_name = type_param.AsValueString() or type_param.AsString()
        except Exception:
            pass

    if not type_name:
        try:
            type_name = getattr(element, "Name", None)
        except Exception:
            type_name = None

    return "{} : {}".format(family_name or "Unknown", type_name or "Unknown")




def _collect_panel_combo_options(groups, panel_lookup):
    options = []
    seen = set()

    def add_option(value):
        name = (value or "").strip()
        if not name:
            return
        key = name.upper()
        if key in seen:
            return
        seen.add(key)
        options.append(name)

    for panel_name in sorted((panel_lookup or {}).keys(), key=lambda v: (v or "").upper()):
        add_option(panel_name)

    for group in groups or []:
        add_option(group.get("panel_name"))
        for member in group.get("members") or []:
            add_option(member.get("panel_name"))

    return options


def _collect_circuit_combo_options(groups):
    options = []
    seen = set()

    def add_option(value):
        name = (value or "").strip()
        if not name:
            return
        key = name.upper()
        if key in seen:
            return
        seen.add(key)
        options.append(name)

    for group in groups or []:
        add_option(group.get("circuit_number"))
        for member in group.get("members") or []:
            add_option(member.get("circuit_number"))

    defaults = [
        "DEDICATED",
        "BYPARENT",
        "SECONDBYPARENT",
        "NONGROUPEDBLOCK",
        "TVTRUSS",
        "EMERGENCY",
        "STANDARD",
    ]
    for value in defaults:
        add_option(value)

    def sort_key(value):
        value = (value or "").strip()
        if re.match(r"^\d+$", value):
            return (0, int(value), value)
        return (1, value.upper(), value)

    return sorted(options, key=sort_key)


def _build_preview_rows(groups, panel_options, circuit_options):
    rows = []
    group_index = 0

    for group in groups or []:
        members = group.get("members") or []
        if not members:
            continue

        if group_index > 0:
            rows.append(
                PreviewRow(
                    group_index=group_index,
                    group_label="",
                    panel_name="",
                    circuit_number="",
                    family_type="",
                    element_id="",
                    source_item=None,
                    row_background=None,
                    row_foreground=DEFAULT_ROW_FOREGROUND,
                    is_spacer=True,
                    panel_options=[],
                    circuit_options=[],
                )
            )

        group_index += 1
        group_size = len(members)
        row_background, row_foreground = _group_row_colors(group_size)
        panel_name = group.get("panel_name") or "NO_PANEL"
        circuit_number = group.get("circuit_number") or "NO_CIRCUIT"
        group_type = (group.get("group_type") or "normal").upper()
        group_label = "{:03d} | {} | PANEL {} | CIRCUIT {} | {} ITEM(S)".format(
            group_index,
            group_type,
            panel_name,
            circuit_number,
            group_size,
        )

        for member in members:
            element = member.get("element")
            elem_id = getattr(getattr(element, "Id", None), "IntegerValue", None)
            row = PreviewRow(
                group_index=group_index,
                group_label=group_label,
                panel_name=member.get("panel_name") or panel_name,
                circuit_number=member.get("circuit_number") or circuit_number,
                family_type=_family_type_text(element),
                element_id=elem_id,
                source_item=member,
                row_background=row_background,
                row_foreground=row_foreground,
                panel_options=panel_options,
                circuit_options=circuit_options,
            )
            rows.append(row)

    return rows

def _show_preview_dialog(groups, panel_lookup):
    panel_options = _collect_panel_combo_options(groups, panel_lookup)
    circuit_options = _collect_circuit_combo_options(groups)
    rows = _build_preview_rows(groups, panel_options, circuit_options)
    if not rows:
        return None

    xaml_path = script.get_bundle_file("SuperCircuitV5Preview.xaml")
    if not xaml_path:
        xaml_path = os.path.join(os.path.dirname(__file__), "SuperCircuitV5Preview.xaml")

    if not os.path.exists(xaml_path):
        forms.alert("Preview UI XAML not found:\n{}".format(xaml_path), title=__title__)
        return None

    window = SuperCircuitPreviewWindow(xaml_path, rows)
    result = window.show_dialog()
    if not result or not window.accepted:
        return False
    return rows


def _build_panel_lookup_upper(panel_lookup):
    upper_lookup = {}
    for name, info in (panel_lookup or {}).items():
        key = (name or "").strip().upper()
        if not key:
            continue
        upper_lookup[key] = (name, info)
    return upper_lookup


def _build_panel_choices(panel_value, upper_lookup):
    choices = []
    seen = set()

    for token in _split_panel_choices(panel_value):
        key = token.strip().upper()
        if not key:
            continue
        panel_entry = upper_lookup.get(key)
        if not panel_entry:
            continue
        canonical_name, panel_info = panel_entry
        panel_elem = panel_info.get("element") if panel_info else None
        panel_id = getattr(getattr(panel_elem, "Id", None), "IntegerValue", None)
        if panel_id is not None and panel_id in seen:
            continue
        if panel_id is not None:
            seen.add(panel_id)
        choices.append(
            {
                "name": canonical_name,
                "element": panel_elem,
                "distribution_system_ids": list((panel_info or {}).get("distribution_system_ids") or []),
            }
        )

    return choices


def _apply_panel_bindings(item, panel_lookup, upper_lookup):
    panel_raw = (item.get("panel_name") or "").strip()
    if not panel_raw:
        item["panel_element"] = None
        item["panel_distribution_system_ids"] = []
        item["panel_choices"] = None
        item["panel_choice_count"] = None
        return

    panel_choices = _build_panel_choices(panel_raw, upper_lookup)
    if panel_choices:
        item["panel_choices"] = panel_choices
        item["panel_choice_count"] = len(panel_choices)
        primary = panel_choices[0]
        item["panel_name"] = primary.get("name")
        item["panel_element"] = primary.get("element")
        item["panel_distribution_system_ids"] = list(primary.get("distribution_system_ids") or [])
        return

    panel_info = panel_lookup.get(panel_raw)
    if not panel_info:
        upper_entry = upper_lookup.get(panel_raw.upper())
        if upper_entry:
            panel_raw = upper_entry[0]
            panel_info = upper_entry[1]
            item["panel_name"] = panel_raw

    if panel_info:
        item["panel_element"] = panel_info.get("element")
        item["panel_distribution_system_ids"] = list(panel_info.get("distribution_system_ids") or [])
        item["panel_choices"] = [
            {
                "name": panel_raw,
                "element": panel_info.get("element"),
                "distribution_system_ids": list(panel_info.get("distribution_system_ids") or []),
            }
        ]
        item["panel_choice_count"] = 1
    else:
        item["panel_element"] = None
        item["panel_distribution_system_ids"] = []
        item["panel_choices"] = None
        choices = _split_panel_choices(panel_raw)
        item["panel_choice_count"] = len(choices) if choices else None


def _apply_preview_edits(rows, panel_lookup):
    if not rows:
        return 0

    upper_lookup = _build_panel_lookup_upper(panel_lookup)
    edited_count = 0

    for row in rows:
        item = getattr(row, "source_item", None)
        if not item:
            continue

        new_panel = (row.panel_name or "").strip()
        new_circuit = (row.circuit_number or "").strip()

        old_panel = (item.get("panel_name") or "").strip()
        old_circuit = (item.get("circuit_number") or "").strip()

        if new_panel == old_panel and new_circuit == old_circuit:
            continue

        item["panel_name"] = new_panel
        item["circuit_number"] = new_circuit
        _apply_panel_bindings(item, panel_lookup, upper_lookup)
        edited_count += 1

    return edited_count


def _update_panel_choice_counts(items):
    for item in items or []:
        panel_choices = item.get("panel_choices")
        if panel_choices:
            item["panel_choice_count"] = len(panel_choices)
            continue

        panel_raw = item.get("panel_name")
        choices = _split_panel_choices(panel_raw)
        if choices:
            unique = {c.strip().upper() for c in choices if c and c.strip()}
            item["panel_choice_count"] = len(unique)
        else:
            item["panel_choice_count"] = None


def _run_client_preprocess(info_items, doc, panel_lookup):
    if client_helpers and hasattr(client_helpers, "preprocess_items"):
        try:
            logger.info("Calling client_helpers.preprocess_items")
            processed = client_helpers.preprocess_items(info_items, doc, panel_lookup, logger)
            if processed:
                info_items = processed
        except Exception as ex:
            logger.warning("Client preprocess_items failed: {}".format(ex))
    return info_items


def main():
    client_key = _select_client()
    if not client_key:
        logger.info("No client selected; aborting.")
        return

    global client_helpers
    client_helpers = _load_client_helpers(client_key)
    if client_helpers:
        logger.info(
            "Client helpers loaded: %s",
            getattr(client_helpers, "__file__", "unknown"),
        )
        logger.info(
            "Client helpers has preprocess_items: %s",
            hasattr(client_helpers, "preprocess_items"),
        )

    selection_only = _select_scope()
    if selection_only is None:
        logger.info("No scope selected; aborting.")
        return

    doc = revit.doc
    panels = list(get_all_panels(doc))
    panel_lookup = circuits.build_panel_lookup(panels)

    elements = _collect_elements(doc, selection_only=selection_only)
    elements = _filter_disallowed_elements(elements)
    if not elements:
        logger.info("No elements found for processing.")
        return

    info_items = circuits.gather_element_info(doc, elements, panel_lookup, logger)
    logger.info("Gathered circuit info items: {}".format(len(info_items)))
    if not info_items:
        logger.info("No elements with circuit data were found.")
        return

    _update_panel_choice_counts(info_items)

    output = script.get_output()
    _debug_ba_da("pre", info_items, panel_lookup, output)
    info_items = _run_client_preprocess(info_items, doc, panel_lookup)
    _update_panel_choice_counts(info_items)
    _debug_ba_da("post", info_items, panel_lookup, output)

    groups = circuits.assemble_groups(info_items, client_helpers, logger)
    if not groups:
        logger.info("Grouping produced no circuit batches.")
        return

    groups = _sort_groups(groups)

    preview_rows = _show_preview_dialog(groups, panel_lookup)
    if preview_rows is False:
        logger.info("Preview cancelled; no circuits created.")
        return
    if preview_rows is None:
        logger.info("Preview could not be shown; no circuits created.")
        return

    edited_count = _apply_preview_edits(preview_rows, panel_lookup)
    if edited_count:
        logger.info("Applied {} edit(s) from preview.".format(edited_count))
        _update_panel_choice_counts(info_items)
        info_items = _run_client_preprocess(info_items, doc, panel_lookup)
        _update_panel_choice_counts(info_items)

        groups = circuits.assemble_groups(info_items, client_helpers, logger)
        if not groups:
            logger.info("Grouping produced no circuit batches after edits.")
            return
        groups = _sort_groups(groups)

    created_systems = _run_creation(doc, groups)
    if not created_systems:
        logger.info("No circuits were created.")
        return

    _run_apply_data(doc, created_systems)

    logger.info("Created {} circuits.".format(len(created_systems)))


if __name__ == "__main__":
    main()















