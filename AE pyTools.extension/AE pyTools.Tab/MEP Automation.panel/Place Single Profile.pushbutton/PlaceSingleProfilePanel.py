# -*- coding: utf-8 -*-
"""
Dockable pane for placing a single profile.
"""

import os
import tempfile
import time
import re
import sys
import ctypes
try:
    from collections.abc import Mapping
except ImportError:
    from collections import Mapping

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import (
    BoundingBoxXYZ,
    BuiltInParameter,
    ElementId,
    FamilySymbol,
    Family,
    FilteredElementCollector,
    ImageExportOptions,
    ExportRange,
    ImageFileType,
    ImageResolution,
    Level,
    RevitLinkInstance,
    FamilyInstance,
    Transaction,
    ViewFamily,
    ViewFamilyType,
    View3D,
    ViewOrientation3D,
    ViewPlan,
    PlanViewRange,
    ViewDetailLevel,
    Options,
    GeometryInstance,
    Solid,
    Curve,
    XYZ,
    ZoomFitType,
)
try:
    from Autodesk.Revit.DB.Structure import StructuralType
except Exception:
    StructuralType = None
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from Autodesk.Revit.UI.Selection import ObjectSnapTypes
from System.Drawing import Size, Color, Bitmap
from System import IntPtr, Uri, TimeSpan
from System.Windows import Int32Rect, Thickness, Point, Rect, Visibility
from System.Windows.Controls import Canvas, Image
from System.Windows.Media import Stretch
from System.Windows.Media import Brushes, DrawingVisual, Pen, StreamGeometry, PixelFormats
from System.Windows.Media.Imaging import BitmapSizeOptions, CroppedBitmap, BitmapImage, BitmapCacheOption, RenderTargetBitmap
from System.Windows.Interop import Imaging
from System.Windows.Shapes import Ellipse, Line, Rectangle
from System.Collections.Generic import List
from System.Windows.Threading import DispatcherTimer

LIB_ROOT = os.path.abspath(
    os.path.join(
        os.path.dirname(__file__),
        "..",
        "..",
        "..",
        "..",
        "CEDLib.lib",
    )
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from LogicClasses.placement_engine import PlaceElementsEngine  # noqa: E402
from LogicClasses.profile_repository import ProfileRepository  # noqa: E402
from LogicClasses.linked_equipment import build_child_requests, find_equipment_by_name  # noqa: E402
from LogicClasses.profile_schema import equipment_defs_to_legacy  # noqa: E402
from ExtensibleStorage.yaml_store import load_active_yaml_data  # noqa: E402


TITLE = "Place Single Profile"
PANEL_ID = "8d3f8f2d-0e6f-4b2d-9e64-8d2a2b57d7b8"
_PREVIEW_FAILED = object()


def _sanitize_equipment_definitions(equipment_defs):
    cleaned_defs = []
    for eq in equipment_defs or []:
        if not isinstance(eq, Mapping):
            continue
        sanitized = dict(eq)
        linked_sets = []
        for linked_set in sanitized.get("linked_sets") or []:
            if not isinstance(linked_set, Mapping):
                continue
            ls_copy = dict(linked_set)
            led_list = []
            for led in ls_copy.get("linked_element_definitions") or []:
                if not isinstance(led, Mapping):
                    continue
                if led.get("is_parent_anchor"):
                    continue
                led_copy = dict(led)
                tags = led_copy.get("tags")
                if isinstance(tags, list):
                    led_copy["tags"] = [t if isinstance(t, Mapping) else {} for t in tags]
                else:
                    led_copy["tags"] = []
                offsets = led_copy.get("offsets")
                if isinstance(offsets, list):
                    led_copy["offsets"] = [o if isinstance(o, Mapping) else {} for o in offsets]
                else:
                    led_copy["offsets"] = [{}]
                led_list.append(led_copy)
            ls_copy["linked_element_definitions"] = led_list
            linked_sets.append(ls_copy)
        sanitized["linked_sets"] = linked_sets
        cleaned_defs.append(sanitized)
    return cleaned_defs


def _sanitize_profiles(profiles):
    cleaned = []
    for prof in profiles or []:
        if not isinstance(prof, Mapping):
            continue
        prof_copy = dict(prof)
        types = []
        for t in prof_copy.get("types") or []:
            if not isinstance(t, Mapping):
                continue
            t_copy = dict(t)
            inst_cfg = t_copy.get("instance_config")
            if not isinstance(inst_cfg, Mapping):
                inst_cfg = {}
            offsets = inst_cfg.get("offsets")
            if not isinstance(offsets, list) or not offsets:
                offsets = [{}]
            inst_cfg["offsets"] = [off if isinstance(off, Mapping) else {} for off in offsets]
            tags = inst_cfg.get("tags")
            if isinstance(tags, list):
                inst_cfg["tags"] = [tag if isinstance(tag, Mapping) else {} for tag in tags]
            else:
                inst_cfg["tags"] = []
            params = inst_cfg.get("parameters")
            if not isinstance(params, Mapping):
                params = {}
            inst_cfg["parameters"] = params
            t_copy["instance_config"] = inst_cfg
            types.append(t_copy)
        prof_copy["types"] = types
        cleaned.append(prof_copy)
    return cleaned


def _is_independent_name(cad_name):
    if not cad_name:
        return False
    trimmed = cad_name.strip()
    if re.match(r"^\d{3}", trimmed):
        return False
    if ":" in trimmed:
        return False
    return not trimmed.lower().startswith("heb")


def _has_parent_relation(raw_data, cad_name):
    target = (cad_name or "").strip().lower()
    if not target:
        return False
    for eq_def in raw_data.get("equipment_definitions") or []:
        eq_name = (eq_def.get("name") or eq_def.get("id") or "").strip().lower()
        eq_id = (eq_def.get("id") or "").strip().lower()
        if target != eq_name and target != eq_id:
            continue
        rel = eq_def.get("linked_relations") or {}
        parent = rel.get("parent") or {}
        parent_id = (parent.get("equipment_id") or "").strip()
        return bool(parent_id)
    return False


def _is_independent_profile(raw_data, cad_name):
    if not _is_independent_name(cad_name):
        return False
    return not _has_parent_relation(raw_data, cad_name)


def _group_truth_profile_choices(raw_data, available_cads, independent_only=False):
    """Collapse equipment definitions by truth-source metadata so only canonical profiles appear."""
    if independent_only:
        available = {(name or "").strip(): True for name in available_cads if _is_independent_profile(raw_data, name)}
    else:
        available = {(name or "").strip(): True for name in available_cads}
    groups = {}
    for eq_def in raw_data.get("equipment_definitions") or []:
        cad_name = (eq_def.get("name") or eq_def.get("id") or "").strip()
        if not cad_name or cad_name not in available:
            continue
        truth_id = (eq_def.get("ced_truth_source_id") or eq_def.get("id") or cad_name).strip()
        if not truth_id:
            truth_id = cad_name
        display_name = (eq_def.get("ced_truth_source_name") or cad_name).strip() or cad_name
        group = groups.setdefault(truth_id, {"display": display_name, "members": [], "primary": None})
        group["members"].append(cad_name)
        eq_id = (eq_def.get("id") or "").strip()
        if eq_id and eq_id == truth_id:
            group["primary"] = cad_name
    if not groups:
        return [{"label": name, "cad": name} for name in sorted(available_cads)]
    display_counts = {}
    for info in groups.values():
        label = info.get("display") or ""
        display_counts[label] = display_counts.get(label, 0) + 1
    options = []
    seen_cads = set()
    for truth_id in sorted(groups.keys()):
        info = groups[truth_id]
        cad = info.get("primary") or (info.get("members") or [None])[0]
        if not cad or cad not in available:
            continue
        label = info.get("display") or cad
        if display_counts.get(label, 0) > 1:
            label = u"{} [{}]".format(label, truth_id)
        options.append({"label": label, "cad": cad})
        seen_cads.add(cad)
    for cad in sorted(available_cads):
        cad = cad.strip()
        if cad and cad not in seen_cads and cad in available:
            options.append({"label": cad, "cad": cad})
    return options


def _normalize_key(value):
    if not value:
        return ""
    return " ".join(str(value).lower().replace("_", " ").split())


def _build_repository(raw_data):
    cleaned_defs = _sanitize_equipment_definitions(raw_data.get("equipment_definitions") or [])
    legacy_profiles = equipment_defs_to_legacy(cleaned_defs)
    cleaned_profiles = _sanitize_profiles(legacy_profiles)
    eq_defs = ProfileRepository._parse_profiles(cleaned_profiles)
    return ProfileRepository(eq_defs)


def _gather_child_requests(parent_def, base_point, base_rotation, repo, data):
    requests = []
    if not parent_def:
        return requests
    for linked_set in parent_def.get("linked_sets") or []:
        for led_entry in linked_set.get("linked_element_definitions") or []:
            led_id = (led_entry.get("id") or "").strip()
            if not led_id:
                continue
            reqs = build_child_requests(repo, data, parent_def, base_point, base_rotation, led_id)
            if reqs:
                requests.extend(reqs)
    return requests


class _PlaceSingleProfileHandler(IExternalEventHandler):
    def __init__(self):
        self._payload = None

    def set_payload(self, raw_data, cad_choice, multi):
        self._payload = {
            "raw_data": raw_data,
            "cad_choice": cad_choice,
            "multi": bool(multi),
        }

    def Execute(self, uiapp):  # noqa: N802
        payload = self._payload
        self._payload = None
        if not payload:
            return
        raw_data = payload.get("raw_data") or {}
        cad_choice = payload.get("cad_choice")
        multi = bool(payload.get("multi"))
        if not cad_choice:
            return
        try:
            doc = uiapp.ActiveUIDocument.Document
            uidoc = uiapp.ActiveUIDocument
            repo = _build_repository(raw_data)
            labels = repo.labels_for_cad(cad_choice)
            if not labels:
                forms.alert("Equipment definition '{}' has no linked types.".format(cad_choice), title=TITLE)
                return
            points = []
            if multi:
                while True:
                    try:
                        base_pt = uidoc.Selection.PickPoint(
                            ObjectSnapTypes.None,
                            "Pick points for '{}' (ESC to finish)".format(cad_choice),
                        )
                    except Exception:
                        break
                    if base_pt is None:
                        break
                    points.append(base_pt)
            else:
                base_pt = uidoc.Selection.PickPoint(
                    ObjectSnapTypes.None,
                    "Pick base point for '{}'".format(cad_choice),
                )
                if base_pt:
                    points = [base_pt]
            if not points:
                return
            selection_map = {cad_choice: labels}
            rows = []
            for base_pt in points:
                rows.append({
                    "Name": cad_choice,
                    "Count": "1",
                    "Position X": str(base_pt.X * 12.0),
                    "Position Y": str(base_pt.Y * 12.0),
                    "Position Z": str(base_pt.Z * 12.0),
                    "Rotation": "0",
                })
            parent_def = find_equipment_by_name(raw_data, cad_choice)
            if parent_def:
                for base_pt in points:
                    child_requests = _gather_child_requests(parent_def, base_pt, 0.0, repo, raw_data)
                    if not child_requests:
                        continue
                    for request in child_requests:
                        name = request.get("name")
                        req_labels = request.get("labels")
                        point = request.get("target_point")
                        rotation = request.get("rotation")
                        if not name or not req_labels or point is None:
                            continue
                        selection_map[name] = req_labels
                        rows.append({
                            "Name": name,
                            "Count": "1",
                            "Position X": str(point.X * 12.0),
                            "Position Y": str(point.Y * 12.0),
                            "Position Z": str(point.Z * 12.0),
                            "Rotation": str(rotation or 0.0),
                        })
            engine = PlaceElementsEngine(doc, repo, allow_tags=False, transaction_name=TITLE)
            results = engine.place_from_csv(rows, selection_map)
            placed = results.get("placed", 0)
            forms.alert("Placed {} element(s) for equipment definition '{}'.".format(placed, cad_choice), title=TITLE)
        except Exception as exc:
            forms.alert("Error during placement:\n\n{}".format(exc), title=TITLE)

    def GetName(self):  # noqa: N802
        return "PlaceSingleProfileHandler"


class _PreviewRenderHandler(IExternalEventHandler):
    def __init__(self):
        self._payload = None

    def set_payload(self, payload=None):
        self._payload = payload

    def Execute(self, uiapp):  # noqa: N802
        self._payload = None
        panel = PlaceSingleProfilePanel.get_instance()
        if panel is None:
            return
        panel._process_preview_queue(uiapp)

    def GetName(self):  # noqa: N802
        return "PreviewRenderHandler"


class PlaceSingleProfilePanel(forms.WPFPanel):
    panel_id = PANEL_ID
    panel_title = "Place Single Profile"
    panel_source = os.path.abspath(os.path.join(os.path.dirname(__file__), "PlaceSingleProfilePanel.xaml"))
    _instance = None

    def __init__(self):
        forms.WPFPanel.__init__(self)
        PlaceSingleProfilePanel._instance = self

        self._raw_data = {}
        self._repo = None
        self._choices = []
        self._choice_map = {}
        self._preview_cache = {}
        self._symbol_cache = {}
        self._symbol_lookup = None
        self._symbol_lookup_info = {}
        self._bbox_cache = {}
        self._preview_reason_cache = {}
        self._preview_zoom = 1.0
        self._preview_last_export_error = None
        self._preview_queue = []
        self._preview_pending = set()
        self._preview_in_handler = False
        self._preview_rendered = 0
        self._preview_total = 0
        self._preview_last_render_ts = None
        self._preview_first_request_ts = None
        self._preview_timer = None
        self._preview_busy = False
        self._preview_ready = False
        self._last_preview_key = None
        self._preview3d_cache = {}
        self._preview3d_pending = set()
        self._preview3d_request = None
        self._preview3d_reason_cache = {}
        self._preview3d_angle = "preview"
        self._preview3d_composite_reason = None

        self._profile_combo = self.FindName("ProfileCombo")
        self._independent_only = self.FindName("IndependentOnlyCheck")
        self._place_button = self.FindName("PlaceButton")
        self._refresh_button = self.FindName("RefreshButton")
        self._preview_canvas = self.FindName("PreviewCanvas")
        self._preview_image = self.FindName("PreviewImage")
        self._preview_zoom_slider = None
        self._preview_zoom_value = None
        self._profile_types_list = self.FindName("ProfileTypesList")
        self._preview3d_image = self.FindName("Preview3DImage")
        self._preview3d_iso = self.FindName("Preview3DIsoButton")
        self._preview3d_front = self.FindName("Preview3DFrontButton")
        self._preview3d_right = self.FindName("Preview3DRightButton")
        self._preview3d_top = self.FindName("Preview3DTopButton")
        self._status_text = self.FindName("StatusText")
        self._place_handler = _PlaceSingleProfileHandler()
        self._place_event = ExternalEvent.Create(self._place_handler)
        self._preview_handler = _PreviewRenderHandler()
        self._preview_event = ExternalEvent.Create(self._preview_handler)
        self._setup_preview_timer()

        if self._independent_only is not None:
            self._independent_only.Checked += self._on_filter_changed
            self._independent_only.Unchecked += self._on_filter_changed
        if self._profile_combo is not None:
            self._profile_combo.SelectionChanged += self._on_profile_changed
            self._profile_combo.Loaded += self._on_profile_loaded
        if self._place_button is not None:
            self._place_button.Click += self._on_place
        self._place_multi_button = self.FindName("PlaceMultiButton")
        if self._place_multi_button is not None:
            self._place_multi_button.Click += self._on_place_multi
        if self._refresh_button is not None:
            self._refresh_button.Click += self._on_refresh
        if self._preview_canvas is not None:
            self._preview_canvas.SizeChanged += self._on_preview_size_changed
            self._preview_canvas.Loaded += self._on_preview_loaded
        self._preview_zoom = 1.0
        if self._preview3d_iso is not None:
            self._preview3d_iso.Click += self._on_preview3d_iso
        if self._preview3d_front is not None:
            self._preview3d_front.Click += self._on_preview3d_front
        if self._preview3d_right is not None:
            self._preview3d_right.Click += self._on_preview3d_right
        if self._preview3d_top is not None:
            self._preview3d_top.Click += self._on_preview3d_top

        self._refresh_data()

    @classmethod
    def get_instance(cls):
        return cls._instance

    def _set_status(self, text):
        if self._status_text is not None:
            self._status_text.Text = text or ""

    def _on_refresh(self, sender, args):
        self._refresh_data()
        # If refresh succeeded but status was left as a load error, overwrite it.
        if self._repo and self._status_text is not None:
            try:
                current = self._status_text.Text or ""
            except Exception:
                current = ""
            if current.startswith("Failed to load active YAML"):
                self._set_status("Loaded {} profiles.".format(len(self._repo.cad_names() or [])))

    def _on_filter_changed(self, sender, args):
        self._refresh_profile_choices()

    def _on_profile_changed(self, sender, args):
        if self._profile_combo is not None:
            try:
                label = self._profile_combo.SelectedItem
            except Exception:
                label = None
            if label:
                try:
                    self._profile_combo.Text = label
                except Exception:
                    pass
        self._symbol_lookup = None
        self._update_preview()

    def _on_profile_loaded(self, sender, args):
        self._update_preview()

    def _on_preview_size_changed(self, sender, args):
        if not self._preview_ready:
            width = getattr(self._preview_canvas, "ActualWidth", 0) if self._preview_canvas is not None else 0
            height = getattr(self._preview_canvas, "ActualHeight", 0) if self._preview_canvas is not None else 0
            if width and height:
                self._preview_ready = True
                self._update_preview()

    def _on_preview_loaded(self, sender, args):
        if self._preview_canvas is None:
            return
        width = getattr(self._preview_canvas, "ActualWidth", 0)
        height = getattr(self._preview_canvas, "ActualHeight", 0)
        if width and height:
            self._preview_ready = True
            self._update_preview()

    def _on_preview_zoom_changed(self, sender, args):
        try:
            self._preview_zoom = float(self._preview_zoom_slider.Value)
        except Exception:
            self._preview_zoom = 1.0
        self._update_zoom_label()
        self._preview_cache = {}
        self._preview_reason_cache = {}
        self._symbol_lookup = None
        self._symbol_lookup_info = {}
        self._preview_queue = []
        self._preview_pending = set()
        self._preview_rendered = 0
        self._preview_total = 0
        self._preview_last_render_ts = None
        self._preview_first_request_ts = None
        self._preview3d_cache = {}
        self._preview3d_pending = set()
        self._preview3d_request = None
        self._preview3d_reason_cache = {}
        self._preview3d_composite_reason = None
        self._update_preview()

    def _update_zoom_label(self):
        if self._preview_zoom_value is None:
            return
        try:
            self._preview_zoom_value.Text = "{:.1f}x".format(self._preview_zoom)
        except Exception:
            self._preview_zoom_value.Text = "1.0x"

    def _setup_preview_timer(self):
        if self._preview_timer is not None:
            return
        try:
            timer = DispatcherTimer()
            timer.Interval = TimeSpan.FromMilliseconds(300)
            timer.Tick += self._on_preview_timer_tick
            timer.Start()
            self._preview_timer = timer
        except Exception:
            self._preview_timer = None

    def _on_preview_timer_tick(self, sender, args):
        if self._preview_in_handler or (not self._preview_queue and not self._preview3d_request and not self._preview3d_pending):
            return
        try:
            self._preview_event.Raise()
        except Exception:
            pass

    def _refresh_data(self):
        try:
            _path, data = load_active_yaml_data()
        except Exception as exc:
            # If we already have data loaded, keep it and just report the load failure.
            if self._raw_data and self._repo:
                self._set_status("Using cached YAML (load failed: {}).".format(exc))
                try:
                    self._refresh_profile_choices()
                except Exception:
                    pass
                return
            self._raw_data = {}
            self._repo = None
            self._choices = []
            self._choice_map = {}
            self._set_status("Failed to load active YAML: {}".format(exc))
            if self._profile_combo is not None:
                self._profile_combo.ItemsSource = []
            self._clear_preview()
            return

        self._raw_data = data
        self._repo = _build_repository(data)
        self._preview_cache = {}
        self._symbol_cache = {}
        self._symbol_lookup = None
        self._symbol_lookup_info = {}
        self._bbox_cache = {}
        self._preview_reason_cache = {}
        self._preview_queue = []
        self._preview_pending = set()
        self._preview_rendered = 0
        self._preview_total = 0
        self._preview_last_render_ts = None
        self._preview_first_request_ts = None
        self._preview3d_cache = {}
        self._preview3d_pending = set()
        self._preview3d_request = None
        self._preview3d_reason_cache = {}
        self._preview3d_composite_reason = None
        self._refresh_profile_choices()

    def _refresh_profile_choices(self):
        if not self._repo:
            return
        cad_names = list(self._repo.cad_names() or [])
        independent_only = bool(self._independent_only.IsChecked) if self._independent_only is not None else False
        raw_choices = _group_truth_profile_choices(self._raw_data, cad_names, independent_only=independent_only)
        raw_choices = sorted(raw_choices, key=lambda entry: (entry.get("label") or "").lower())
        self._choice_map = {entry.get("label"): entry.get("cad") for entry in raw_choices}
        labels = [entry.get("label") for entry in raw_choices]
        if self._profile_combo is not None:
            self._profile_combo.ItemsSource = labels
            if labels:
                self._profile_combo.SelectedIndex = 0
                self._profile_combo.Text = labels[0]
                self._set_status("Loaded {} profiles.".format(len(labels)))
            else:
                self._set_status("No profiles available.")
        self._update_preview()

    def _selected_profile(self):
        if self._profile_combo is None:
            return None
        label = self._profile_combo.SelectedItem
        if label is None:
            label = getattr(self._profile_combo, "Text", None)
        if not label:
            return None
        return self._choice_map.get(label, label)

    def _clear_preview(self):
        if self._preview_canvas is not None:
            self._preview_canvas.Children.Clear()
        if self._preview_image is not None:
            self._preview_image.Source = None
        self._last_preview_key = None

    def _collect_preview_points(self, parent_def, cad_choice=None):
        points = []
        if not parent_def:
            parent_def = find_equipment_by_name(self._raw_data, cad_choice) if cad_choice else None
        for linked_set in parent_def.get("linked_sets") or []:
            for led_entry in linked_set.get("linked_element_definitions") or []:
                label = led_entry.get("label")
                for offset in led_entry.get("offsets") or []:
                    try:
                        x = float(offset.get("x_inches", 0.0))
                    except Exception:
                        x = 0.0
                    try:
                        y = float(offset.get("y_inches", 0.0))
                    except Exception:
                        y = 0.0
                    points.append({"x": x, "y": y, "label": label})
        if points:
            return points
        if self._repo and cad_choice:
            try:
                labels = list(self._repo.labels_for_cad(cad_choice) or [])
            except Exception:
                labels = []
            for label in labels:
                try:
                    linked_def = self._repo.definition_for_label(cad_choice, label)
                except Exception:
                    linked_def = None
                if not linked_def:
                    continue
                placement = linked_def.get_placement()
                if not placement:
                    continue
                offset = placement.get_offset_xyz()
                if offset:
                    points.append({"x": offset[0] * 12.0, "y": offset[1] * 12.0, "label": label})
        return points

    def _update_preview(self):
        if self._preview_canvas is None and self._preview_image is None:
            return
        if self._preview_busy:
            return
        self._preview_busy = True
        try:
            self._set_status("Preview updating...")
            self._clear_preview()
            cad_choice = self._selected_profile()
            if not cad_choice:
                self._update_profile_type_list([])
                self._set_status("Select a profile to preview.")
                return
            labels = []
            if self._repo:
                try:
                    labels = list(self._repo.labels_for_cad(cad_choice) or [])
                except Exception:
                    labels = []
            self._update_profile_type_list(labels)
            preview_key = None
            if labels:
                preview_key = labels[0]
                if preview_key != self._last_preview_key:
                    self._last_preview_key = preview_key
                cached = self._preview_cache.get(preview_key)
                if cached is None:
                    cached = self._label_preview_source(preview_key)
                if cached is not None and self._preview_image is not None:
                    self._preview_image.Source = cached
            parent_def = find_equipment_by_name(self._raw_data, cad_choice)
            entries = self._collect_preview_points(parent_def, cad_choice=cad_choice)
            if not entries:
                self._set_status("No child offsets to preview for '{}'.".format(cad_choice))
                entries = [{"x": 0.0, "y": 0.0, "label": None}]

            for entry in entries:
                label = entry.get("label")
                bbox = self._symbol_bbox_inches(label)
                if bbox:
                    width_in, height_in = bbox
                else:
                    width_in = 12.0
                    height_in = 12.0
                entry["width_in"] = max(width_in, 1.0)
                entry["height_in"] = max(height_in, 1.0)
                entry["preview"] = self._label_preview_source(label)

            min_x = min(entry["x"] - entry["width_in"] / 2.0 for entry in entries)
            max_x = max(entry["x"] + entry["width_in"] / 2.0 for entry in entries)
            min_y = min(entry["y"] - entry["height_in"] / 2.0 for entry in entries)
            max_y = max(entry["y"] + entry["height_in"] / 2.0 for entry in entries)
            span_x = max(max_x - min_x, 1.0)
            span_y = max(max_y - min_y, 1.0)

            if labels:
                self._update_3d_composite_preview(labels, entries, min_x, max_x, min_y, max_y)

            canvas_w = (self._preview_canvas.ActualWidth or self._preview_canvas.Width or 260.0)
            canvas_h = (self._preview_canvas.ActualHeight or self._preview_canvas.Height or 240.0)
            margin = 18.0
            scale = min((canvas_w - 2 * margin) / span_x, (canvas_h - 2 * margin) / span_y)
            if scale <= 0:
                scale = 1.0

            def map_point(pt):
                x, y = pt
                px = (x - min_x) * scale + margin
                py = (max_y - y) * scale + margin
                return px, py

            border = Rectangle()
            border.Width = max(canvas_w - 2, 10)
            border.Height = max(canvas_h - 2, 10)
            border.Stroke = Brushes.DimGray
            border.StrokeThickness = 1
            Canvas.SetLeft(border, 1)
            Canvas.SetTop(border, 1)
            self._preview_canvas.Children.Add(border)

            missing_previews = []
            pending_previews = []
            missing_reasons = []
            for entry in entries:
                px, py = map_point((entry["x"], entry["y"]))
                rect_w = max(entry["width_in"] * scale, 6.0)
                rect_h = max(entry["height_in"] * scale, 6.0)
                preview = entry.get("preview")
                if preview is not None:
                    bg_rect = Rectangle()
                    bg_rect.Width = rect_w
                    bg_rect.Height = rect_h
                    bg_rect.Fill = Brushes.DimGray
                    bg_rect.StrokeThickness = 0
                    Canvas.SetLeft(bg_rect, px - rect_w / 2.0)
                    Canvas.SetTop(bg_rect, py - rect_h / 2.0)
                    self._preview_canvas.Children.Add(bg_rect)
                if preview is not None:
                    img = Image()
                    img.Source = preview
                    img.Width = rect_w
                    img.Height = rect_h
                    img.Stretch = Stretch.UniformToFill
                    Canvas.SetLeft(img, px - rect_w / 2.0)
                    Canvas.SetTop(img, py - rect_h / 2.0)
                    self._preview_canvas.Children.Add(img)
                else:
                    label = entry.get("label") or "<unknown>"
                    if label in self._preview_pending:
                        if label not in pending_previews:
                            pending_previews.append(label)
                    else:
                        if label not in missing_previews:
                            missing_previews.append(label)
                            reason = self._preview_reason_cache.get(label)
                            if reason:
                                missing_reasons.append("{} ({})".format(label, reason))
                    self._draw_receptacle_marker(px, py, rect_w, rect_h)
                if preview is not None:
                    rect = Rectangle()
                    rect.Width = rect_w
                    rect.Height = rect_h
                    rect.Stroke = Brushes.DeepSkyBlue
                    rect.StrokeThickness = 0.8
                    rect.Fill = Brushes.Transparent
                    Canvas.SetLeft(rect, px - rect_w / 2.0)
                    Canvas.SetTop(rect, py - rect_h / 2.0)
                    self._preview_canvas.Children.Add(rect)

            if len(entries) > 1:
                mid_x = (min_x + max_x) / 2.0
                mid_y = (min_y + max_y) / 2.0
                origin_px, origin_py = map_point((mid_x, mid_y))
                x_axis = Line()
                x_axis.X1 = origin_px - 20
                x_axis.X2 = origin_px + 20
                x_axis.Y1 = origin_py
                x_axis.Y2 = origin_py
                x_axis.Stroke = Brushes.Gray
                x_axis.StrokeThickness = 0.5
                self._preview_canvas.Children.Add(x_axis)
                y_axis = Line()
                y_axis.X1 = origin_px
                y_axis.X2 = origin_px
                y_axis.Y1 = origin_py - 20
                y_axis.Y2 = origin_py + 20
                y_axis.Stroke = Brushes.Gray
                y_axis.StrokeThickness = 0.5
                self._preview_canvas.Children.Add(y_axis)
            else:
                px, py = map_point((entries[0]["x"], entries[0]["y"]))
                cross = Line()
                cross.X1 = px - 6
                cross.X2 = px + 6
                cross.Y1 = py
                cross.Y2 = py
                cross.Stroke = Brushes.Gray
                cross.StrokeThickness = 0.5
                self._preview_canvas.Children.Add(cross)
                cross_v = Line()
                cross_v.X1 = px
                cross_v.X2 = px
                cross_v.Y1 = py - 6
                cross_v.Y2 = py + 6
                cross_v.Stroke = Brushes.Gray
                cross_v.StrokeThickness = 0.5
                self._preview_canvas.Children.Add(cross_v)
            if missing_reasons:
                self._set_status("Preview render failed: {}".format(", ".join(missing_reasons)))
            elif pending_previews:
                stalled = False
                try:
                    if self._preview_last_render_ts is not None:
                        stalled = (time.time() - self._preview_last_render_ts) > 5.0
                    elif self._preview_first_request_ts is not None:
                        stalled = (time.time() - self._preview_first_request_ts) > 5.0
                except Exception:
                    stalled = False
                if stalled:
                    self._set_status("Preview render stalled (queue={}, pending={}, 3d_pending={}).".format(
                        len(self._preview_queue),
                        len(self._preview_pending),
                        len(self._preview3d_pending),
                    ))
                else:
                    self._set_status("Rendering previews for: {} ({} pending)".format(", ".join(pending_previews[:3]), len(self._preview_pending)))
            elif missing_previews:
                detailed = []
                for label in missing_previews:
                    reason = self._preview_reason_cache.get(label) or "Unknown failure"
                    detailed.append("{} ({})".format(label, reason))
                self._set_status("Preview render failed: {}".format(", ".join(detailed)))
            elif self._preview3d_composite_reason:
                self._set_status("3D preview failed: {}".format(self._preview3d_composite_reason))
            elif preview_key:
                key3d = (preview_key, self._preview3d_angle)
                if self._preview3d_cache.get(key3d) is _PREVIEW_FAILED:
                    reason = self._preview3d_reason_cache.get(key3d) or "Unknown failure"
                    self._set_status("3D preview failed: {} ({})".format(preview_key, reason))
                elif self._preview3d_pending:
                    self._set_status("Rendering 3D previews... ({} pending)".format(len(self._preview3d_pending)))
                else:
                    try:
                        total = len(entries)
                        with_images = len([entry for entry in entries if entry.get("preview") is not None])
                        self._set_status("Preview ready: {} item(s), {} image(s).".format(total, with_images))
                    except Exception:
                        self._set_status("Preview ready.")
            elif self._preview3d_pending:
                self._set_status("Rendering 3D previews... ({} pending)".format(len(self._preview3d_pending)))
            else:
                try:
                    total = len(entries)
                    with_images = len([entry for entry in entries if entry.get("preview") is not None])
                    self._set_status("Preview ready: {} item(s), {} image(s).".format(total, with_images))
                except Exception:
                    self._set_status("Preview ready.")
        except Exception as exc:
            self._set_status("Preview error: {}".format(exc))
        finally:
            self._preview_busy = False

    def _update_profile_type_list(self, labels):
        if self._profile_types_list is None:
            return
        if not labels:
            try:
                self._profile_types_list.ItemsSource = []
            except Exception:
                pass
            return
        items = [label for label in labels if label]
        try:
            self._profile_types_list.ItemsSource = items
        except Exception:
            pass

    def _symbol_for_label(self, doc, label):
        if not label or doc is None:
            return None
        cached = self._symbol_cache.get(label)
        if cached is not None:
            return cached
        symbol = self._find_symbol_for_label(doc, label)
        if symbol is not None:
            self._symbol_cache[label] = symbol
        return symbol

    def _symbol_bbox_inches(self, label):
        if not label:
            return None
        cached = self._bbox_cache.get(label)
        if cached is not None:
            return cached
        symbol = self._symbol_for_label(revit.doc, label)
        if symbol is None:
            return None
        try:
            bbox = symbol.get_BoundingBox(None)
        except Exception:
            bbox = None
        if bbox is None:
            return None
        try:
            width = abs(bbox.Max.X - bbox.Min.X) * 12.0
            height = abs(bbox.Max.Y - bbox.Min.Y) * 12.0
        except Exception:
            return None
        if width <= 1e-6 or height <= 1e-6:
            return None
        self._bbox_cache[label] = (width, height)
        return self._bbox_cache[label]

    def _label_preview_source(self, label):
        if not label:
            return None
        if label in self._preview_cache:
            cached = self._preview_cache.get(label)
            if cached is _PREVIEW_FAILED:
                return None
            if cached is not None:
                return cached
        self._request_preview_render(label)
        return None

    def _preview3d_preview_for_label(self, label):
        if not label:
            return None
        key = (label, "preview")
        cached = self._preview3d_cache.get(key)
        if cached is _PREVIEW_FAILED:
            return None
        if cached is not None:
            return cached
        try:
            symbol = self._symbol_for_label(revit.doc, label)
        except Exception:
            symbol = None
        if symbol is None:
            self._preview3d_cache[key] = _PREVIEW_FAILED
            self._preview3d_reason_cache[key] = "Family type not loaded."
            return None
        preview = self._symbol_preview_source(symbol)
        if preview is None:
            self._preview3d_cache[key] = _PREVIEW_FAILED
            self._preview3d_reason_cache[key] = "Preview image missing."
            return None
        self._preview3d_cache[key] = preview
        return preview

    def _update_3d_composite_preview(self, labels, entries, min_x, max_x, min_y, max_y):
        if self._preview3d_image is None:
            return
        self._preview3d_composite_reason = None
        if not labels or not entries:
            self._preview3d_image.Source = None
            self._preview3d_composite_reason = "No items to preview."
            return
        width = self._preview3d_image.ActualWidth or self._preview3d_image.Width or 320.0
        height = self._preview3d_image.ActualHeight or self._preview3d_image.Height or 200.0
        width = max(width, 160.0)
        height = max(height, 120.0)
        margin = 10.0
        span_x = max(max_x - min_x, 1.0)
        span_y = max(max_y - min_y, 1.0)
        scale = min((width - 2 * margin) / span_x, (height - 2 * margin) / span_y)
        if scale <= 0:
            scale = 1.0

        def map_point(pt):
            x, y = pt
            px = (x - min_x) * scale + margin
            py = (max_y - y) * scale + margin
            return px, py

        visual = DrawingVisual()
        dc = visual.RenderOpen()
        dc.DrawRectangle(Brushes.White, None, Rect(0, 0, width, height))
        drawn = 0
        for entry in entries:
            label = entry.get("label")
            preview = self._preview3d_preview_for_label(label)
            if preview is None:
                continue
            px, py = map_point((entry["x"], entry["y"]))
            rect_w = max(entry["width_in"] * scale, 16.0)
            rect_h = max(entry["height_in"] * scale, 16.0)
            rect = Rect(px - rect_w / 2.0, py - rect_h / 2.0, rect_w, rect_h)
            dc.DrawImage(preview, rect)
            drawn += 1
        dc.Close()
        if drawn == 0:
            self._preview3d_image.Source = None
            self._preview3d_composite_reason = "No preview images available."
            return
        bmp = RenderTargetBitmap(int(width), int(height), 96, 96, PixelFormats.Pbgra32)
        bmp.Render(visual)
        try:
            bmp.Freeze()
        except Exception:
            pass
        self._preview3d_image.Source = bmp

    def _update_3d_preview_for_label(self, label):
        if self._preview3d_image is None:
            return
        if not label:
            self._preview3d_image.Source = None
            return
        key = (label, self._preview3d_angle)
        cached = self._preview3d_cache.get(key)
        if cached is _PREVIEW_FAILED:
            self._preview3d_image.Source = None
            reason = self._preview3d_reason_cache.get(key)
            if reason:
                self._set_status("3D preview failed: {} ({})".format(label, reason))
            return
        if cached is not None:
            self._preview3d_image.Source = cached
            return
        try:
            symbol = self._symbol_for_label(revit.doc, label)
        except Exception:
            symbol = None
        if symbol is None:
            self._preview3d_cache[key] = _PREVIEW_FAILED
            self._preview3d_reason_cache[key] = "Family type not loaded."
            self._preview3d_image.Source = None
            self._set_status("3D preview failed: {} (Family type not loaded.)".format(label))
            return
        preview = self._symbol_preview_source(symbol)
        if preview is None:
            self._preview3d_cache[key] = _PREVIEW_FAILED
            self._preview3d_reason_cache[key] = "Preview image missing."
            self._preview3d_image.Source = None
            self._set_status("3D preview failed: {} (Preview image missing.)".format(label))
            return
        self._preview3d_cache[key] = preview
        self._preview3d_image.Source = preview

    def _request_3d_render(self, label, angle):
        if not label or self._preview_in_handler:
            return
        key = (label, angle)
        if key in self._preview3d_pending:
            return
        self._preview3d_pending.add(key)
        self._preview3d_request = key
        try:
            self._preview_event.Raise()
        except Exception:
            pass

    def _on_preview3d_iso(self, sender, args):
        self._preview3d_angle = "iso"
        self._update_3d_preview_for_label(self._last_preview_key)

    def _on_preview3d_front(self, sender, args):
        self._preview3d_angle = "front"
        self._update_3d_preview_for_label(self._last_preview_key)

    def _on_preview3d_right(self, sender, args):
        self._preview3d_angle = "right"
        self._update_3d_preview_for_label(self._last_preview_key)

    def _on_preview3d_top(self, sender, args):
        self._preview3d_angle = "top"
        self._update_3d_preview_for_label(self._last_preview_key)

    def _selected_profile_label(self):
        if self._profile_combo is None:
            return None
        label = self._profile_combo.SelectedItem
        if label is None:
            label = getattr(self._profile_combo, "Text", None)
        if not label:
            return None
        return label

    def _first_label_with_preview(self, labels):
        if not labels:
            return None
        for label in labels:
            try:
                symbol = self._symbol_for_label(revit.doc, label)
            except Exception:
                symbol = None
            if symbol is None:
                continue
            preview = self._symbol_preview_source(symbol)
            if preview is not None:
                return label
        return None

    def _draw_receptacle_marker(self, px, py, rect_w, rect_h):
        if self._preview_canvas is None:
            return
        radius = min(rect_w, rect_h) / 2.0
        if radius < 6.0:
            radius = 6.0
        ellipse = Ellipse()
        ellipse.Width = radius * 2.0
        ellipse.Height = radius * 2.0
        ellipse.Stroke = Brushes.LimeGreen
        ellipse.StrokeThickness = 1.0
        ellipse.Fill = Brushes.Transparent
        Canvas.SetLeft(ellipse, px - radius)
        Canvas.SetTop(ellipse, py - radius)
        self._preview_canvas.Children.Add(ellipse)

        line_len = radius * 1.2
        stem = Line()
        stem.X1 = px
        stem.X2 = px
        stem.Y1 = py + radius * 0.2
        stem.Y2 = py + line_len
        stem.Stroke = Brushes.LimeGreen
        stem.StrokeThickness = 1.0
        self._preview_canvas.Children.Add(stem)

        body_w = radius * 0.9
        body_h = radius * 0.45
        body = Rectangle()
        body.Width = body_w
        body.Height = body_h
        body.Stroke = Brushes.LimeGreen
        body.StrokeThickness = 1.0
        body.Fill = Brushes.Transparent
        Canvas.SetLeft(body, px - body_w / 2.0)
        Canvas.SetTop(body, py + line_len)
        self._preview_canvas.Children.Add(body)

    def _request_preview_render(self, label):
        if not label or self._preview_in_handler:
            return
        if label in self._preview_cache:
            return
        if label in self._preview_pending:
            return
        self._preview_pending.add(label)
        self._preview_queue.append(label)
        if self._preview_first_request_ts is None:
            try:
                self._preview_first_request_ts = time.time()
            except Exception:
                self._preview_first_request_ts = None
        if self._preview_total == 0:
            self._preview_total = len(self._preview_pending)
        try:
            self._preview_event.Raise()
        except Exception:
            pass

    def _process_preview_queue(self, uiapp=None):
        if self._preview_in_handler:
            return
        self._preview_in_handler = True
        try:
            max_per_pass = 1
            processed = 0
            while self._preview_queue and processed < max_per_pass:
                label = self._preview_queue.pop(0)
                if not label:
                    continue
                try:
                    symbol = self._symbol_for_label(revit.doc, label)
                    rendered, reason = self._render_plan_preview(symbol, uiapp=uiapp)
                    cached = rendered
                    if cached is None and not reason:
                        cached = self._symbol_preview_source(symbol)
                    if cached is None and reason:
                        if reason == "Family type not loaded.":
                            detail = self._symbol_lookup_hint(label)
                            if detail:
                                reason = "{} {}".format(reason, detail)
                        self._preview_reason_cache[label] = reason
                    if cached is None and not reason:
                        self._preview_reason_cache[label] = "Preview image generation failed."
                    self._preview_cache[label] = cached if cached is not None else _PREVIEW_FAILED
                except Exception as exc:
                    self._preview_reason_cache[label] = "Preview error: {}".format(exc)
                    self._preview_cache[label] = _PREVIEW_FAILED
                if label in self._preview_pending:
                    self._preview_pending.remove(label)
                processed += 1
                self._preview_rendered += 1
                try:
                    self._preview_last_render_ts = time.time()
                except Exception:
                    self._preview_last_render_ts = None
            if self._preview3d_request:
                label, angle = self._preview3d_request
                self._preview3d_request = None
                try:
                    symbol = self._symbol_for_label(revit.doc, label)
                    rendered, reason = self._render_3d_preview(symbol, angle, uiapp=uiapp)
                    key = (label, angle)
                    if rendered is None and reason:
                        self._preview_reason_cache[label] = reason
                    if rendered is None and reason:
                        self._preview3d_reason_cache[key] = reason
                    self._preview3d_cache[key] = rendered if rendered is not None else _PREVIEW_FAILED
                except Exception as exc:
                    self._preview3d_cache[(label, angle)] = _PREVIEW_FAILED
                    self._preview3d_reason_cache[(label, angle)] = "3D preview error: {}".format(exc)
                if (label, angle) in self._preview3d_pending:
                    self._preview3d_pending.remove((label, angle))
                try:
                    self._update_3d_preview_for_label(label)
                except Exception:
                    pass
        finally:
            self._preview_in_handler = False
        try:
            if self._preview_queue:
                self._set_status("Rendering previews... {}/{}".format(self._preview_rendered, max(self._preview_total, 1)))
                try:
                    self._preview_event.Raise()
                except Exception:
                    pass
            self._update_preview()
        except Exception:
            pass

    def _symbol_lookup_hint(self, label):
        try:
            lookup = self._symbol_label_lookup(revit.doc)
        except Exception:
            return ""
        if not lookup:
            try:
                doc_title = getattr(revit.doc, "Title", "<doc>")
                open_docs = list(revit.doc.Application.Documents)
                open_titles = [getattr(d, "Title", "<doc>") for d in open_docs if d is not None]
                counts = []
                for d in open_docs:
                    if d is None:
                        continue
                    try:
                        fam_count = FilteredElementCollector(d).OfClass(Family).WhereElementIsNotElementType().GetElementCount()
                    except Exception:
                        fam_count = -1
                    try:
                        sym_count = FilteredElementCollector(d).OfClass(FamilySymbol).WhereElementIsElementType().GetElementCount()
                    except Exception:
                        sym_count = -1
                    try:
                        inst_count = FilteredElementCollector(d).OfClass(FamilyInstance).WhereElementIsNotElementType().GetElementCount()
                    except Exception:
                        inst_count = -1
                    counts.append("{}(F={},S={},I={})".format(getattr(d, "Title", "<doc>"), fam_count, sym_count, inst_count))
            except Exception:
                doc_title = getattr(revit.doc, "Title", "<doc>")
                open_titles = []
                counts = []
            info = self._symbol_lookup_info or {}
            sym_count = info.get("symbol_count", 0)
            label_count = info.get("label_count", 0)
            samples = info.get("samples") or []
            sample_text = ", ".join(samples[:3]) if samples else "none"
            return "(no symbols found in doc '{}', open docs={}, counts={}, sym_count={}, label_count={}, samples={})".format(
                doc_title,
                ", ".join(open_titles),
                "; ".join(counts),
                sym_count,
                label_count,
                sample_text,
            )
        norm_label = _normalize_key(label)
        fam_name, type_name = self._split_label(label)
        type_norm = _normalize_key(type_name) if type_name else ""
        fam_norm = _normalize_key(fam_name) if fam_name else ""
        matches = []
        if type_norm:
            for key, sym in lookup.items():
                if key.endswith(type_norm):
                    matches.append(sym)
        if matches:
            return "(found {} symbols matching type name, label='{}')".format(len(matches), norm_label)
        if fam_norm:
            for key, sym in lookup.items():
                if key.startswith(fam_norm):
                    matches.append(sym)
            if matches:
                return "(found {} symbols matching family name, label='{}')".format(len(matches), norm_label)
        return "(label='{}', lookup keys={})".format(norm_label, len(lookup))

    def _render_plan_preview(self, symbol, uiapp=None):
        if symbol is None:
            return None, "Family type not loaded."
        doc = revit.doc
        try:
            if symbol.Document is not None and symbol.Document != doc:
                return None, "Symbol is in a linked document."
        except Exception:
            pass
        level = self._get_first_level(doc)
        vft = self._get_floorplan_view_family_type(doc)
        if level is None or vft is None:
            return None, "No floor plan view type or level."
        temp_view = None
        temp_inst = None
        try:
            if uiapp is None or uiapp.ActiveUIDocument is None:
                return None, "No active UIDocument."
        except Exception:
            return None, "No active UIDocument."
        try:
            if uiapp.ActiveUIDocument.Document != doc:
                return None, "Active UIDocument mismatch."
        except Exception:
            return None, "Active UIDocument mismatch."
        try:
            t = Transaction(doc, "CED Temp Preview")
            t.Start()
            try:
                if not symbol.IsActive:
                    symbol.Activate()
            except Exception:
                pass
            try:
                temp_view = ViewPlan.Create(doc, vft.Id, level.Id)
            except Exception:
                t.RollBack()
                return None, "Failed to create temp view."
            try:
                temp_view.Name = "CED Preview {}".format(int(time.time()))
            except Exception:
                pass
            try:
                if StructuralType is not None:
                    temp_inst = doc.Create.NewFamilyInstance(XYZ(0, 0, 0), symbol, level, StructuralType.NonStructural)
                else:
                    temp_inst = doc.Create.NewFamilyInstance(XYZ(0, 0, 0), symbol, level)
            except Exception:
                t.RollBack()
                return None, "Failed to place temp instance (hosted family?)."
            try:
                doc.Regenerate()
            except Exception:
                pass
            # Force plan-style graphics.
            try:
                temp_view.DetailLevel = ViewDetailLevel.Fine
            except Exception:
                pass
            try:
                # Avoid hard dependency on ViewDisplayStyle for older API versions.
                from Autodesk.Revit.DB import ViewDisplayStyle  # noqa: F401
                try:
                    temp_view.DisplayStyle = ViewDisplayStyle.Wireframe
                except Exception:
                    temp_view.DisplayStyle = ViewDisplayStyle.HiddenLine
            except Exception:
                pass
            try:
                # Force a typical plan view range so symbolic plan graphics show.
                view_range = temp_view.GetViewRange()
                if view_range is None:
                    view_range = PlanViewRange()
                # offsets are in feet
                view_range.SetOffset(PlanViewRange.TopClipPlane, 10.0)
                view_range.SetOffset(PlanViewRange.CutPlane, 4.0)
                view_range.SetOffset(PlanViewRange.BottomClipPlane, 0.0)
                view_range.SetOffset(PlanViewRange.ViewDepthPlane, 0.0)
                temp_view.SetViewRange(view_range)
            except Exception:
                pass
            try:
                temp_view.AreAnnotationCategoriesHidden = True
            except Exception:
                pass
            try:
                temp_view.AreAnalyticalModelCategoriesHidden = True
            except Exception:
                pass
            try:
                temp_view.AreImportCategoriesHidden = True
            except Exception:
                pass
            try:
                temp_view.AreCoordinationModelCategoriesHidden = True
            except Exception:
                pass
            try:
                temp_view.AreModelCategoriesHidden = False
            except Exception:
                pass
            bbox = None
            try:
                bbox = temp_inst.get_BoundingBox(temp_view)
            except Exception:
                bbox = None
            if bbox is None:
                try:
                    bbox = temp_inst.get_BoundingBox(None)
                except Exception:
                    bbox = None
            if bbox is not None:
                try:
                    zoom = self._preview_zoom if self._preview_zoom else 1.0
                    if zoom < 0.1:
                        zoom = 0.1
                    scale = 1.0 / zoom
                    dx = max(bbox.Max.X - bbox.Min.X, 0.01)
                    dy = max(bbox.Max.Y - bbox.Min.Y, 0.01)
                    dz = max(bbox.Max.Z - bbox.Min.Z, 0.01)
                    max_dim = max(dx, dy, dz)
                    pad = max(max_dim * 0.1, 1.0)
                    cx = (bbox.Min.X + bbox.Max.X) * 0.5
                    cy = (bbox.Min.Y + bbox.Max.Y) * 0.5
                    cz = (bbox.Min.Z + bbox.Max.Z) * 0.5
                    half_x = dx * 0.5 * scale + pad
                    half_y = dy * 0.5 * scale + pad
                    half_z = dz * 0.5 + pad
                    min_pt = XYZ(cx - half_x, cy - half_y, cz - half_z)
                    max_pt = XYZ(cx + half_x, cy + half_y, cz + half_z)
                    crop = BoundingBoxXYZ()
                    crop.Min = min_pt
                    crop.Max = max_pt
                    temp_view.CropBox = crop
                    temp_view.CropBoxActive = True
                    temp_view.CropBoxVisible = False
                except Exception:
                    pass
            try:
                temp_view.Scale = 100
            except Exception:
                pass
            t.Commit()
            vector_img, vector_reason = self._render_plan_vector_preview(temp_inst, temp_view, bbox=bbox)
            if vector_img is not None:
                return vector_img, None
            if vector_reason:
                return None, "Plan geometry render failed. {}".format(vector_reason)
            img_path = self._export_view_image(doc, temp_view)
            if img_path:
                if self._is_blank_bitmap(img_path):
                    return None, "Image export blank."
                return self._load_bitmap_image(img_path), None
            export_err = self._preview_last_export_error
            if export_err:
                return None, "Image export failed. {}".format(export_err)
            return None, "Image export failed."
        finally:
            t_cleanup = None
            try:
                t_cleanup = Transaction(doc, "CED Temp Preview Cleanup")
                t_cleanup.Start()
                ids = List[ElementId]()
                if temp_inst is not None:
                    ids.Add(temp_inst.Id)
                if temp_view is not None:
                    ids.Add(temp_view.Id)
                if ids.Count > 0:
                    doc.Delete(ids)
                t_cleanup.Commit()
            except Exception:
                try:
                    if t_cleanup is not None:
                        t_cleanup.RollBack()
                except Exception:
                    pass
        return None, "Unknown render failure."

    def _render_3d_preview(self, symbol, angle, uiapp=None):
        if symbol is None:
            return None, "Family type not loaded."
        doc = revit.doc
        try:
            if uiapp is None or uiapp.ActiveUIDocument is None:
                return None, "No active UIDocument."
        except Exception:
            return None, "No active UIDocument."
        vft = self._get_3d_view_family_type(doc)
        level = self._get_first_level(doc)
        if vft is None or level is None:
            return None, "No 3D view type or level."
        temp_view = None
        temp_inst = None
        try:
            t = Transaction(doc, "CED Temp 3D Preview")
            t.Start()
            try:
                if not symbol.IsActive:
                    symbol.Activate()
            except Exception:
                pass
            try:
                temp_view = View3D.CreateIsometric(doc, vft.Id)
            except Exception:
                t.RollBack()
                return None, "Failed to create 3D view."
            try:
                if StructuralType is not None:
                    temp_inst = doc.Create.NewFamilyInstance(XYZ(0, 0, 0), symbol, level, StructuralType.NonStructural)
                else:
                    temp_inst = doc.Create.NewFamilyInstance(XYZ(0, 0, 0), symbol, level)
            except Exception:
                t.RollBack()
                return None, "Failed to place temp instance."
            try:
                doc.Regenerate()
            except Exception:
                pass
            bbox = None
            try:
                bbox = temp_inst.get_BoundingBox(None)
            except Exception:
                bbox = None
            size = None
            center = None
            if bbox is not None:
                try:
                    dx = max(bbox.Max.X - bbox.Min.X, 0.01)
                    dy = max(bbox.Max.Y - bbox.Min.Y, 0.01)
                    dz = max(bbox.Max.Z - bbox.Min.Z, 0.01)
                    size = max(dx, dy, dz)
                    center = XYZ(
                        (bbox.Min.X + bbox.Max.X) * 0.5,
                        (bbox.Min.Y + bbox.Max.Y) * 0.5,
                        (bbox.Min.Z + bbox.Max.Z) * 0.5,
                    )
                    expand = max(size * 0.2, 1.0)
                    min_pt = XYZ(bbox.Min.X - expand, bbox.Min.Y - expand, bbox.Min.Z - expand)
                    max_pt = XYZ(bbox.Max.X + expand, bbox.Max.Y + expand, bbox.Max.Z + expand)
                    crop = BoundingBoxXYZ()
                    crop.Min = min_pt
                    crop.Max = max_pt
                    temp_view.SetSectionBox(crop)
                    try:
                        temp_view.IsSectionBoxActive = True
                    except Exception:
                        pass
                except Exception:
                    pass
            try:
                self._apply_3d_orientation(temp_view, angle, center=center, size=size)
            except Exception:
                pass
            t.Commit()
            img_path = self._export_view_image(doc, temp_view)
            if img_path:
                if self._is_blank_bitmap(img_path):
                    return None, "3D export blank."
                return self._load_bitmap_image(img_path), None
            export_err = self._preview_last_export_error
            if export_err:
                return None, "3D export failed. {}".format(export_err)
            return None, "3D export failed."
        finally:
            t_cleanup = None
            try:
                t_cleanup = Transaction(doc, "CED Temp 3D Preview Cleanup")
                t_cleanup.Start()
                ids = List[ElementId]()
                if temp_inst is not None:
                    ids.Add(temp_inst.Id)
                if temp_view is not None:
                    ids.Add(temp_view.Id)
                if ids.Count > 0:
                    doc.Delete(ids)
                t_cleanup.Commit()
            except Exception:
                try:
                    if t_cleanup is not None:
                        t_cleanup.RollBack()
                except Exception:
                    pass
        return None, "Unknown 3D render failure."

    def _render_plan_vector_preview(self, instance, view, bbox=None):
        if instance is None or view is None:
            return None, "Missing instance or view."
        geom = None
        try:
            options = Options()
            options.View = view
            options.DetailLevel = ViewDetailLevel.Fine
            try:
                options.IncludeNonVisibleObjects = True
            except Exception:
                pass
            geom = instance.get_Geometry(options)
        except Exception as exc:
            geom = None

        polylines = []
        curve_polylines = []

        def add_curve(curve, bucket=None):
            if curve is None:
                return
            pts = None
            try:
                pts = curve.Tessellate()
            except Exception:
                pts = None
            if not pts:
                try:
                    pts = [curve.GetEndPoint(0), curve.GetEndPoint(1)]
                except Exception:
                    pts = None
            if pts and len(pts) >= 2:
                if bucket is None:
                    polylines.append(pts)
                else:
                    bucket.append(pts)

        def walk_geom(obj):
            if obj is None:
                return
            if isinstance(obj, GeometryInstance):
                try:
                    inst_geom = obj.GetInstanceGeometry()
                except Exception:
                    inst_geom = None
                if inst_geom:
                    for g in inst_geom:
                        walk_geom(g)
                return
            if isinstance(obj, Solid):
                try:
                    for edge in obj.Edges:
                        try:
                            add_curve(edge.AsCurve())
                        except Exception:
                            pass
                except Exception:
                    pass
                return
            if isinstance(obj, Curve):
                add_curve(obj, curve_polylines)
                return

        if geom:
            try:
                for g in geom:
                    walk_geom(g)
            except Exception:
                pass

        if curve_polylines:
            polylines = curve_polylines
        if not polylines and bbox is not None:
            try:
                min_x = bbox.Min.X
                max_x = bbox.Max.X
                min_y = bbox.Min.Y
                max_y = bbox.Max.Y
                if max_x > min_x and max_y > min_y:
                    pts = [
                        XYZ(min_x, min_y, 0),
                        XYZ(max_x, min_y, 0),
                        XYZ(max_x, max_y, 0),
                        XYZ(min_x, max_y, 0),
                        XYZ(min_x, min_y, 0),
                    ]
                    polylines.append(pts)
            except Exception:
                pass

        if not polylines:
            return None, "No plan curves found."

        pts2d = []
        for poly in polylines:
            for p in poly:
                try:
                    pts2d.append((p.X, p.Y))
                except Exception:
                    pass
        if not pts2d:
            return None, "No plan points found."

        min_x = min(pt[0] for pt in pts2d)
        max_x = max(pt[0] for pt in pts2d)
        min_y = min(pt[1] for pt in pts2d)
        max_y = max(pt[1] for pt in pts2d)
        span_x = max(max_x - min_x, 0.01)
        span_y = max(max_y - min_y, 0.01)

        size = 256.0
        zoom = self._preview_zoom if self._preview_zoom else 1.0
        if zoom < 0.1:
            zoom = 0.1
        pad = size * 0.1 / zoom
        if pad < 4.0:
            pad = 4.0
        if pad > size * 0.45:
            pad = size * 0.45
        scale = min((size - 2.0 * pad) / span_x, (size - 2.0 * pad) / span_y)
        if scale <= 0:
            scale = 1.0

        def map_point(p):
            x, y = p
            px = (x - min_x) * scale + pad
            py = (max_y - y) * scale + pad
            return px, py

        visual = DrawingVisual()
        dc = visual.RenderOpen()
        dc.DrawRectangle(Brushes.White, None, Rect(0, 0, size, size))
        pen = Pen(Brushes.Black, 1.0)
        for poly in polylines:
            if len(poly) < 2:
                continue
            geom_path = StreamGeometry()
            ctx = geom_path.Open()
            try:
                p0 = map_point((poly[0].X, poly[0].Y))
                ctx.BeginFigure(Point(p0[0], p0[1]), False, False)
                for pt in poly[1:]:
                    p = map_point((pt.X, pt.Y))
                    ctx.LineTo(Point(p[0], p[1]), True, False)
            finally:
                ctx.Close()
            try:
                geom_path.Freeze()
            except Exception:
                pass
            dc.DrawGeometry(None, pen, geom_path)
        dc.Close()
        bmp = RenderTargetBitmap(int(size), int(size), 96, 96, PixelFormats.Pbgra32)
        bmp.Render(visual)
        try:
            bmp.Freeze()
        except Exception:
            pass
        return bmp, None

    def _apply_3d_orientation(self, view, angle, center=None, size=None):
        if view is None:
            return
        if angle == "front":
            eye = XYZ(0, -1, 0)
            forward = XYZ(0, 1, 0)
            up = XYZ(0, 0, 1)
        elif angle == "right":
            eye = XYZ(1, 0, 0)
            forward = XYZ(-1, 0, 0)
            up = XYZ(0, 0, 1)
        elif angle == "top":
            eye = XYZ(0, 0, 1)
            forward = XYZ(0, 0, -1)
            up = XYZ(0, 1, 0)
        else:
            eye = XYZ(1, -1, 1)
            forward = XYZ(-1, 1, -1)
            up = XYZ(0, 0, 1)
        try:
            if center is not None and size is not None:
                try:
                    direction = forward.Normalize()
                    distance = max(size * 2.5, 5.0)
                    eye = center + direction.Multiply(distance)
                    forward = (center - eye).Normalize()
                except Exception:
                    pass
            view.SetOrientation(ViewOrientation3D(eye, up, forward))
        except Exception:
            pass

    def _export_view_image(self, doc, view):
        if doc is None or view is None:
            return None
        self._preview_last_export_error = None
        view_name = None
        view_type = None
        view_id = None
        is_template = None
        can_print = None
        original_view = None
        ui_doc = None
        try:
            view_name = getattr(view, "Name", None)
            view_type = getattr(view, "ViewType", None)
            view_id = getattr(view, "Id", None)
            is_template = getattr(view, "IsTemplate", None)
            can_print = getattr(view, "CanBePrinted", None)
            try:
                ui_doc = revit.uidoc
                if ui_doc is not None:
                    original_view = ui_doc.ActiveView
            except Exception:
                ui_doc = None
                original_view = None
        except Exception:
            pass
        temp_dir = tempfile.gettempdir()
        stamp = "{}_{}".format(view.Id.IntegerValue, int(time.time() * 1000))
        export_root = os.path.join(temp_dir, "ced_preview_exports", stamp)
        try:
            if not os.path.exists(export_root):
                os.makedirs(export_root)
        except Exception:
            export_root = temp_dir
        prefix = os.path.join(export_root, "preview")
        options = ImageExportOptions()
        options.FilePath = prefix
        options.ExportRange = ExportRange.SetOfViews
        try:
            options.SetViewsAndSheets(List[ElementId]([view.Id]))
        except Exception:
            self._preview_last_export_error = "Failed to set views for export. view='{}' id={} type={} template={} printable={} dir='{}'".format(
                view_name,
                view_id,
                view_type,
                is_template,
                can_print,
                export_root,
            )
            return None
        try:
            if hasattr(options, "ImageFileType"):
                options.ImageFileType = ImageFileType.PNG
            elif hasattr(options, "FileType"):
                options.FileType = ImageFileType.PNG
        except Exception:
            pass
        try:
            if hasattr(options, "ImageResolution"):
                options.ImageResolution = ImageResolution.DPI_150
        except Exception:
            pass
        try:
            options.ZoomType = ZoomFitType.FitToPage
        except Exception:
            pass
        try:
            doc.ExportImage(options)
        except Exception as exc:
            self._preview_last_export_error = "ExportImage error: {}".format(exc)
            return None
        candidates = []
        try:
            for name in os.listdir(export_root):
                if name.lower().endswith(".png"):
                    candidates.append(os.path.join(export_root, name))
        except Exception:
            self._preview_last_export_error = "ExportImage produced no files (failed to scan export dir). dir='{}'".format(export_root)
            return None
        if not candidates:
            # Fallback: try exporting current view by temporarily activating the temp view.
            try:
                if ui_doc is not None:
                    ui_doc.ActiveView = view
                    options.ExportRange = ExportRange.CurrentView
                    try:
                        doc.ExportImage(options)
                    except Exception as exc:
                        self._preview_last_export_error = "ExportImage current view error: {}".format(exc)
                    try:
                        if original_view is not None:
                            ui_doc.ActiveView = original_view
                    except Exception:
                        pass
                    candidates = []
                    try:
                        for name in os.listdir(export_root):
                            if name.lower().endswith(".png"):
                                candidates.append(os.path.join(export_root, name))
                    except Exception:
                        candidates = []
            except Exception:
                pass
        if not candidates:
            self._preview_last_export_error = "ExportImage produced no files. dir='{}' prefix='{}' view='{}' id={} type={} template={} printable={}".format(
                export_root,
                prefix,
                view_name,
                view_id,
                view_type,
                is_template,
                can_print,
            )
            return None
        candidates.sort(key=lambda path: os.path.getmtime(path))
        return candidates[-1]

    def _load_bitmap_image(self, path):
        if not path or not os.path.exists(path):
            return None
        try:
            bmp = BitmapImage()
            bmp.BeginInit()
            bmp.UriSource = Uri(path)
            bmp.CacheOption = BitmapCacheOption.OnLoad
            bmp.EndInit()
            bmp.Freeze()
            return bmp
        except Exception:
            return None
        finally:
            try:
                os.remove(path)
            except Exception:
                pass

    def _is_blank_bitmap(self, path):
        if not path or not os.path.exists(path):
            return True
        try:
            bmp = Bitmap(path)
        except Exception:
            return True
        try:
            width = bmp.Width
            height = bmp.Height
            if width <= 2 or height <= 2:
                return True
            step_x = max(width // 10, 1)
            step_y = max(height // 10, 1)
            min_bright = 255
            max_bright = 0
            for x in range(0, width, step_x):
                for y in range(0, height, step_y):
                    try:
                        color = bmp.GetPixel(x, y)
                    except Exception:
                        continue
                    bright = (int(color.R) + int(color.G) + int(color.B)) / 3.0
                    if bright < min_bright:
                        min_bright = bright
                    if bright > max_bright:
                        max_bright = bright
            if (max_bright - min_bright) < 8:
                return True
            return False
        except Exception:
            return True
        finally:
            try:
                bmp.Dispose()
            except Exception:
                pass

    def _get_floorplan_view_family_type(self, doc):
        if doc is None:
            return None
        try:
            types = list(FilteredElementCollector(doc).OfClass(ViewFamilyType))
        except Exception:
            types = []
        for vft in types:
            try:
                if vft.ViewFamily == ViewFamily.FloorPlan:
                    return vft
            except Exception:
                continue
        return None

    def _get_3d_view_family_type(self, doc):
        if doc is None:
            return None
        try:
            types = list(FilteredElementCollector(doc).OfClass(ViewFamilyType))
        except Exception:
            types = []
        for vft in types:
            try:
                if vft.ViewFamily == ViewFamily.ThreeDimensional:
                    return vft
            except Exception:
                continue
        return None

    def _get_first_level(self, doc):
        if doc is None:
            return None
        try:
            levels = list(FilteredElementCollector(doc).OfClass(Level))
        except Exception:
            levels = []
        if not levels:
            return None
        levels.sort(key=lambda lvl: getattr(lvl, "Elevation", 0.0))
        return levels[0]

    def _split_label(self, label):
        if not label:
            return None, None
        if ":" not in label:
            return label.strip(), ""
        fam, typ = label.split(":", 1)
        return fam.strip(), typ.strip()

    def _find_symbol_for_label(self, doc, label):
        lookup = self._symbol_label_lookup(doc)
        if not lookup:
            return None
        label_key = _normalize_key(label)
        symbol = lookup.get(label_key)
        if symbol is not None:
            return symbol
        fam_name, type_name = self._split_label(label)
        if fam_name and type_name:
            fam_norm = _normalize_key(fam_name)
            type_norm = _normalize_key(type_name)
            alt_key = _normalize_key(u"{} : {}".format(fam_name, type_name))
            symbol = lookup.get(alt_key)
            if symbol is not None:
                return symbol
            # Fallback: unique type name match
            matches = []
            for key, sym in lookup.items():
                if key.endswith(type_norm):
                    matches.append(sym)
            if len(matches) == 1:
                return matches[0]
        return None

    def _symbol_label_lookup(self, doc):
        if self._symbol_lookup is not None and len(self._symbol_lookup) > 0:
            return self._symbol_lookup
        self._symbol_lookup = None
        self._symbol_lookup_info = {}
        lookup = {}
        symbols = []
        try:
            symbols = list(FilteredElementCollector(doc).OfClass(FamilySymbol).WhereElementIsElementType())
        except Exception:
            symbols = []
        if not symbols:
            try:
                families = list(FilteredElementCollector(doc).OfClass(Family))
            except Exception:
                families = []
            for fam in families:
                try:
                    for sym_id in fam.GetFamilySymbolIds():
                        try:
                            sym = doc.GetElement(sym_id)
                        except Exception:
                            sym = None
                        if sym is not None:
                            symbols.append(sym)
                except Exception:
                    continue
        if not symbols:
            # Try any other open documents in this Revit session.
            try:
                app_docs = list(revit.doc.Application.Documents)
            except Exception:
                app_docs = []
            for other_doc in app_docs:
                try:
                    if other_doc is None or other_doc.Equals(doc):
                        continue
                except Exception:
                    continue
                try:
                    symbols.extend(list(FilteredElementCollector(other_doc).OfClass(FamilySymbol).WhereElementIsElementType()))
                except Exception:
                    continue
        if not symbols:
            try:
                links = list(FilteredElementCollector(doc).OfClass(RevitLinkInstance))
            except Exception:
                links = []
            for link in links:
                try:
                    link_doc = link.GetLinkDocument()
                except Exception:
                    link_doc = None
                if link_doc is None:
                    continue
                try:
                    symbols.extend(list(FilteredElementCollector(link_doc).OfClass(FamilySymbol).WhereElementIsElementType()))
                except Exception:
                    continue
        if not symbols:
            # Final fallback: gather symbols from instances in active + open + linked docs.
            candidate_docs = [doc]
            try:
                candidate_docs.extend([d for d in revit.doc.Application.Documents if d is not None])
            except Exception:
                pass
            for cdoc in candidate_docs:
                try:
                    instances = list(FilteredElementCollector(cdoc).OfClass(FamilyInstance).WhereElementIsNotElementType())
                except Exception:
                    instances = []
                for inst in instances:
                    try:
                        sym = inst.Symbol
                    except Exception:
                        sym = None
                    if sym is not None:
                        symbols.append(sym)
        for sym in symbols:
            sym_fam, sym_type = self._symbol_names(sym)
            if not sym_fam or not sym_type:
                continue
            label_a = u"{} : {}".format(sym_fam, sym_type)
            label_b = u"{}:{}".format(sym_fam, sym_type)
            label_c = u"{} {}".format(sym_fam, sym_type)
            for label_key in (label_a, label_b, label_c, sym_type, sym_fam):
                key = _normalize_key(label_key)
                if key and key not in lookup:
                    lookup[key] = sym
        samples = []
        for sym in symbols[:5]:
            sym_fam, sym_type = self._symbol_names(sym)
            if sym_fam or sym_type:
                samples.append(u"{} : {}".format(sym_fam or "?", sym_type or "?"))
        self._symbol_lookup_info = {
            "symbol_count": len(symbols),
            "label_count": len(lookup),
            "samples": samples,
        }
        self._symbol_lookup = lookup
        return lookup

    def _symbol_names(self, sym):
        sym_fam = None
        sym_type = None
        try:
            fam = sym.Family
            sym_fam = fam.Name if fam else None
        except Exception:
            sym_fam = None
        try:
            sym_type = sym.Name
        except Exception:
            sym_type = None
        if not sym_type:
            try:
                param = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if param:
                    sym_type = param.AsString()
            except Exception:
                pass
        if not sym_fam:
            try:
                param = sym.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                if param:
                    sym_fam = param.AsString()
            except Exception:
                pass
        if not sym_fam:
            try:
                param = sym.get_Parameter(BuiltInParameter.FAMILY_NAME)
                if param:
                    sym_fam = param.AsString()
            except Exception:
                pass
        if not sym_fam:
            try:
                param = sym.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME)
                if param:
                    sym_fam = param.AsString()
            except Exception:
                pass
        if sym_fam:
            sym_fam = sym_fam.strip()
        if sym_type:
            sym_type = sym_type.strip()
        return sym_fam, sym_type

    def _symbol_preview_source(self, symbol, size_px=512):
        if symbol is None:
            return None
        try:
            img = symbol.GetPreviewImage(Size(size_px, size_px))
        except Exception:
            return None
        if img is None:
            return None
        try:
            hbitmap = img.GetHbitmap()
        except Exception:
            return None
        try:
            source = Imaging.CreateBitmapSourceFromHBitmap(
                hbitmap,
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions(),
            )
            cropped = self._crop_preview_image(img, source)
            cropped.Freeze()
            return cropped
        finally:
            try:
                handle = hbitmap.ToInt64() if hasattr(hbitmap, "ToInt64") else int(hbitmap)
                ctypes.windll.gdi32.DeleteObject(handle)
            except Exception:
                pass

    def _crop_preview_image(self, img, source):
        try:
            width = img.Width
            height = img.Height
        except Exception:
            return source
        if width <= 1 or height <= 1:
            return source
        try:
            bg = self._estimate_background(img, width, height)
        except Exception:
            return source
        min_x = width
        min_y = height
        max_x = -1
        max_y = -1
        for y in range(height):
            for x in range(width):
                try:
                    color = img.GetPixel(x, y)
                except Exception:
                    continue
                if self._is_foreground_pixel(color, bg):
                    if x < min_x:
                        min_x = x
                    if y < min_y:
                        min_y = y
                    if x > max_x:
                        max_x = x
                    if y > max_y:
                        max_y = y
        if max_x < min_x or max_y < min_y:
            return self._force_center_crop(source, width, height)
        pad = 2
        min_x = max(min_x - pad, 0)
        min_y = max(min_y - pad, 0)
        max_x = min(max_x + pad, width - 1)
        max_y = min(max_y + pad, height - 1)
        try:
            rect_w = max_x - min_x + 1
            rect_h = max_y - min_y + 1
            rect = self._apply_zoom_rect(min_x, min_y, rect_w, rect_h, width, height)
        except Exception:
            return source
        try:
            return CroppedBitmap(source, rect)
        except Exception:
            return source

    def _is_foreground_pixel(self, color, bg):
        try:
            if color.A == 0:
                return False
        except Exception:
            pass
        dist = self._color_distance(color, bg)
        if dist > 20:
            return True
        try:
            bright = (int(color.R) + int(color.G) + int(color.B)) / 3.0
            bg_bright = (int(bg.R) + int(bg.G) + int(bg.B)) / 3.0
        except Exception:
            return False
        if bg_bright > 160:
            return bright < (bg_bright - 25)
        if bg_bright < 95:
            return bright > (bg_bright + 25)
        return abs(bg_bright - bright) > 25

    def _color_distance(self, a, b):
        try:
            dr = int(a.R) - int(b.R)
            dg = int(a.G) - int(b.G)
            db = int(a.B) - int(b.B)
        except Exception:
            return 0
        return abs(dr) + abs(dg) + abs(db)

    def _estimate_background(self, img, width, height):
        samples = {}
        step = max(min(width, height) // 20, 4)
        for x in range(0, width, step):
            for y in (0, height - 1):
                self._tally_bg_sample(samples, img.GetPixel(x, y))
        for y in range(0, height, step):
            for x in (0, width - 1):
                self._tally_bg_sample(samples, img.GetPixel(x, y))
        if not samples:
            return img.GetPixel(0, 0)
        best = None
        best_count = -1
        for key, value in samples.items():
            if value > best_count:
                best_count = value
                best = key
        r, g, b = best
        return Color.FromArgb(255, r, g, b)

    def _tally_bg_sample(self, samples, color):
        try:
            r = int(color.R) // 16 * 16
            g = int(color.G) // 16 * 16
            b = int(color.B) // 16 * 16
        except Exception:
            return
        key = (r, g, b)
        samples[key] = samples.get(key, 0) + 1

    def _force_center_crop(self, source, width, height):
        try:
            crop_w = int(width * 0.8)
            crop_h = int(height * 0.8)
            min_x = max((width - crop_w) // 2, 0)
            min_y = max((height - crop_h) // 2, 0)
            rect = self._apply_zoom_rect(min_x, min_y, crop_w, crop_h, width, height)
        except Exception:
            return source
        try:
            return CroppedBitmap(source, rect)
        except Exception:
            return source

    def _apply_zoom_rect(self, min_x, min_y, rect_w, rect_h, width, height):
        try:
            zoom = float(self._preview_zoom or 1.0)
        except Exception:
            zoom = 1.0
        if zoom <= 0:
            zoom = 1.0
        if abs(zoom - 1.0) <= 0.001:
            return Int32Rect(min_x, min_y, rect_w, rect_h)
        try:
            center_x = min_x + rect_w / 2.0
            center_y = min_y + rect_h / 2.0
            new_w = int(max(rect_w / zoom, 1))
            new_h = int(max(rect_h / zoom, 1))
            new_min_x = int(max(center_x - new_w / 2.0, 0))
            new_min_y = int(max(center_y - new_h / 2.0, 0))
            if new_min_x + new_w > width:
                new_min_x = max(width - new_w, 0)
            if new_min_y + new_h > height:
                new_min_y = max(height - new_h, 0)
            return Int32Rect(new_min_x, new_min_y, new_w, new_h)
        except Exception:
            return Int32Rect(min_x, min_y, rect_w, rect_h)

    def _on_place(self, sender, args):
        self._place_at_points(single=True)

    def _on_place_multi(self, sender, args):
        try:
            forms.alert(
                "Pick the points you want to place this profile, after you have picked all points, press ESC to finish and see the profiles place.",
                title="Pick Points and Place",
                warn_icon=False,
            )
        except Exception:
            pass
        self._place_at_points(single=False)

    def _place_at_points(self, single=True):
        if not self._repo:
            self._set_status("No active YAML loaded.")
            return
        cad_choice = self._selected_profile()
        if not cad_choice:
            self._set_status("Select a profile to place.")
            return
        self._place_handler.set_payload(self._raw_data, cad_choice, multi=not single)
        self._place_event.Raise()
        self._set_status("Placement requested for '{}'.".format(cad_choice))


def ensure_panel_visible():
    try:
        forms.register_dockable_panel(PlaceSingleProfilePanel, default_visible=False)
    except Exception:
        pass
    try:
        forms.open_dockable_panel(PlaceSingleProfilePanel.panel_id)
    except Exception:
        try:
            forms.open_dockable_panel(PlaceSingleProfilePanel)
        except Exception:
            pass
    panel = PlaceSingleProfilePanel.get_instance()
    if panel:
        try:
            panel.Visibility = Visibility.Visible
        except Exception:
            pass
        try:
            panel.Show()
        except Exception:
            pass
        try:
            panel.BringIntoView()
        except Exception:
            pass
        panel._refresh_data()
