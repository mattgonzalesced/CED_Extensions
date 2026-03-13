# -*- coding: utf-8 -*-
"""
Dockable pane for placing a single profile.
"""

import os
import re
import sys
try:
    from collections.abc import Mapping
except ImportError:
    from collections import Mapping

from pyrevit import forms, revit
from Autodesk.Revit.DB import XYZ
try:
    from Autodesk.Revit.DB.Structure import StructuralType
except Exception:
    StructuralType = None
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from Autodesk.Revit.UI.Selection import ObjectSnapTypes, ObjectType
from System.Windows import Visibility

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


def _mapping_list(value):
    if not isinstance(value, list):
        return []
    return [item for item in value if isinstance(item, Mapping)]


def _clean_display_text(value):
    if value in (None, ""):
        return ""
    try:
        text = str(value)
    except Exception:
        return ""
    return " ".join(text.split())


def _trim_text(value, max_len=48):
    text = _clean_display_text(value)
    if not text:
        return ""
    if len(text) <= max_len:
        return text
    return "{}...".format(text[: max_len - 3].rstrip())


def _extract_keynote_value(entry):
    if not isinstance(entry, Mapping):
        return ""
    key_value = entry.get("key_value")
    if key_value not in (None, ""):
        return _clean_display_text(key_value)
    params = entry.get("parameters")
    if not isinstance(params, Mapping):
        return ""
    for key, value in params.items():
        if (key or "").strip().lower() in ("keynote value", "key value", "keynote"):
            return _clean_display_text(value)
    return ""


def _tag_preview_name(tag_data):
    if not isinstance(tag_data, Mapping):
        return ""
    family = _clean_display_text(tag_data.get("family_name") or tag_data.get("family"))
    type_name = _clean_display_text(tag_data.get("type_name") or tag_data.get("type"))
    if family and type_name:
        return "{} : {}".format(family, type_name)
    if type_name:
        return type_name
    if family:
        return family
    category = _clean_display_text(tag_data.get("category_name") or tag_data.get("category"))
    return category


def _text_note_preview_name(note_data):
    if not isinstance(note_data, Mapping):
        return ""
    note_text = _trim_text(note_data.get("text"), max_len=36)
    note_type = _clean_display_text(note_data.get("type_name"))
    if note_text and note_type:
        return '"{}" ({})'.format(note_text, note_type)
    if note_text:
        return '"{}"'.format(note_text)
    if note_type:
        return note_type
    return ""


def _keynote_preview_name(keynote_data):
    if not isinstance(keynote_data, Mapping):
        return ""
    key_value = _extract_keynote_value(keynote_data)
    type_name = _clean_display_text(keynote_data.get("type_name") or keynote_data.get("type"))
    family = _clean_display_text(keynote_data.get("family_name") or keynote_data.get("family"))
    if key_value and type_name:
        return "{} ({})".format(key_value, type_name)
    if key_value:
        return key_value
    if type_name and family:
        return "{} : {}".format(family, type_name)
    if type_name:
        return type_name
    if family:
        return family
    return ""


def _preview_list(items, name_getter, empty_text="None", max_items=2):
    names = []
    seen = set()
    for item in items or []:
        name = name_getter(item)
        if not name or name in seen:
            continue
        names.append(name)
        seen.add(name)
    if not names:
        return empty_text
    if len(names) <= max_items:
        return ", ".join(names)
    return "{}, {} (+{} more)".format(names[0], names[1], len(names) - 2)


def _cleaned_profiles_from_raw(raw_data):
    cleaned_defs = _sanitize_equipment_definitions(raw_data.get("equipment_definitions") or [])
    legacy_profiles = equipment_defs_to_legacy(cleaned_defs)
    return _sanitize_profiles(legacy_profiles)


def _build_profile_type_assets(cleaned_profiles):
    """
    Build per-CAD lookup keyed by the same unique labels emitted by ProfileRepository.
    """
    assets_by_cad = {}
    for profile in cleaned_profiles or []:
        if not isinstance(profile, Mapping):
            continue
        cad_name = (profile.get("cad_name") or profile.get("equipment_def_id") or "").strip()
        if not cad_name:
            continue
        label_assets = {}
        for type_entry in profile.get("types") or []:
            if not isinstance(type_entry, Mapping):
                continue
            label = (type_entry.get("label") or "").strip()
            if not label:
                continue
            inst_cfg = type_entry.get("instance_config")
            if not isinstance(inst_cfg, Mapping):
                inst_cfg = {}
            offsets = inst_cfg.get("offsets")
            if not isinstance(offsets, list) or not offsets:
                offsets = [{}]
            shared_assets = {
                "tags": _mapping_list(inst_cfg.get("tags")),
                "keynotes": _mapping_list(inst_cfg.get("keynotes")),
                "text_notes": _mapping_list(inst_cfg.get("text_notes")),
            }
            for idx in range(len(offsets)):
                base_label = label if idx == 0 else u"{} #{}".format(label, idx + 1)
                unique_label = base_label
                suffix = 2
                while unique_label in label_assets:
                    unique_label = u"{} #{}".format(base_label, suffix)
                    suffix += 1
                label_assets[unique_label] = shared_assets
        assets_by_cad[cad_name] = label_assets
    return assets_by_cad


def _build_repository_from_profiles(cleaned_profiles):
    eq_defs = ProfileRepository._parse_profiles(cleaned_profiles)
    return ProfileRepository(eq_defs)


def _format_profile_type_item(label, assets):
    assets = assets or {}
    tags = assets.get("tags") or []
    text_notes = assets.get("text_notes") or []
    keynotes = assets.get("keynotes") or []
    tag_preview = _preview_list(tags, _tag_preview_name)
    note_preview = _preview_list(text_notes, _text_note_preview_name)
    keynote_preview = _preview_list(keynotes, _keynote_preview_name)
    return (
        u"{label}\n"
        u"  Tags ({tag_count}): {tag_preview}\n"
        u"  Text Notes ({note_count}): {note_preview}\n"
        u"  Keynotes ({keynote_count}): {keynote_preview}"
    ).format(
        label=label,
        tag_count=len(tags),
        tag_preview=tag_preview,
        note_count=len(text_notes),
        note_preview=note_preview,
        keynote_count=len(keynotes),
        keynote_preview=keynote_preview,
    )


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
    cleaned_profiles = _cleaned_profiles_from_raw(raw_data)
    return _build_repository_from_profiles(cleaned_profiles)


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


def _reference_point_from_element(elem):
    if elem is None:
        return None
    try:
        loc = elem.Location
    except Exception:
        loc = None
    if loc is not None and hasattr(loc, "Point"):
        try:
            return loc.Point
        except Exception:
            pass
    if loc is not None and hasattr(loc, "Curve"):
        try:
            curve = loc.Curve
        except Exception:
            curve = None
        if curve is not None:
            try:
                return curve.Evaluate(0.5, True)
            except Exception:
                pass
    try:
        bbox = elem.get_BoundingBox(None)
    except Exception:
        bbox = None
    if not bbox:
        try:
            bbox = elem.get_BoundingBox(revit.active_view)
        except Exception:
            bbox = None
    if not bbox:
        return None
    return (bbox.Min + bbox.Max) * 0.5


def _collect_reference_points(uidoc):
    if uidoc is None:
        return []
    doc = uidoc.Document
    points = []
    elem_ids = []
    try:
        elem_ids = list(uidoc.Selection.GetElementIds())
    except Exception:
        elem_ids = []
    if not elem_ids:
        try:
            refs = uidoc.Selection.PickObjects(
                ObjectType.Element,
                "Select reference elements (ESC to finish)",
            )
        except Exception:
            refs = None
        if refs:
            elem_ids = [r.ElementId for r in refs if r is not None]
    for elem_id in elem_ids:
        try:
            elem = doc.GetElement(elem_id)
        except Exception:
            elem = None
        pt = _reference_point_from_element(elem)
        if pt is not None:
            points.append(pt)
    return points


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


class _PlaceOnReferenceHandler(IExternalEventHandler):
    def __init__(self):
        self._payload = None

    def set_payload(self, raw_data, cad_choice):
        self._payload = {
            "raw_data": raw_data,
            "cad_choice": cad_choice,
        }

    def Execute(self, uiapp):  # noqa: N802
        payload = self._payload
        self._payload = None
        if not payload:
            return
        raw_data = payload.get("raw_data") or {}
        cad_choice = payload.get("cad_choice")
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
            points = _collect_reference_points(uidoc)
            if not points:
                forms.alert("No reference elements selected.", title=TITLE)
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
        return "PlaceOnReferenceHandler"


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
        self._choice_map = {}
        self._profile_type_assets = {}
        self._symbol_cache = {}
        self._symbol_lookup = None
        self._symbol_lookup_info = {}
        self._active_doc_identity = None
        self._data_doc_identity = None

        self._profile_combo = self.FindName("ProfileCombo")
        self._active_doc_text = self.FindName("ActiveDocText")
        self._independent_only = self.FindName("IndependentOnlyCheck")
        self._place_button = self.FindName("PlaceButton")
        self._refresh_button = self.FindName("RefreshButton")
        self._place_on_reference_button = self.FindName("PlaceOnReferenceButton")
        self._profile_types_list = self.FindName("ProfileTypesList")
        self._status_text = self.FindName("StatusText")
        self._place_handler = _PlaceSingleProfileHandler()
        self._place_event = ExternalEvent.Create(self._place_handler)
        self._place_on_reference_handler = _PlaceOnReferenceHandler()
        self._place_on_reference_event = ExternalEvent.Create(self._place_on_reference_handler)

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
        if self._place_on_reference_button is not None:
            self._place_on_reference_button.Click += self._on_place_on_reference
        if self._refresh_button is not None:
            self._refresh_button.Click += self._on_refresh
        self._sync_active_document(force=True)

    @classmethod
    def get_instance(cls):
        return cls._instance

    def _set_status(self, text):
        if self._status_text is not None:
            self._status_text.Text = text or ""

    def _get_active_doc(self):
        try:
            return getattr(revit, "doc", None)
        except Exception:
            return None

    def _doc_identity(self, doc):
        if doc is None:
            return "<none>"
        try:
            path = doc.PathName or ""
        except Exception:
            path = ""
        try:
            title = doc.Title or "<untitled>"
        except Exception:
            title = "<untitled>"
        return "{}|{}".format(path, title)

    def _doc_display(self, doc):
        if doc is None:
            return "<no active document>"
        try:
            title = doc.Title or "<untitled>"
        except Exception:
            title = "<untitled>"
        try:
            path = doc.PathName or ""
        except Exception:
            path = ""
        return "{} ({})".format(title, path) if path else title

    def _update_active_doc_text(self, doc=None):
        if self._active_doc_text is None:
            return
        if doc is None:
            doc = self._get_active_doc()
        self._active_doc_text.Text = "Active document (ES source): {}".format(self._doc_display(doc))

    def _sync_active_document(self, force=False):
        doc = self._get_active_doc()
        identity = self._doc_identity(doc)
        changed = identity != self._active_doc_identity
        if not force and not changed:
            return False
        self._active_doc_identity = identity
        self._update_active_doc_text(doc)
        self._refresh_data(doc=doc, doc_switched=changed)
        return True

    def _on_refresh(self, sender, args):
        self._update_active_doc_text()
        self._refresh_data(doc=self._get_active_doc(), doc_switched=False)
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
        self._update_profile_type_list_for_selection()

    def _on_profile_loaded(self, sender, args):
        self._update_profile_type_list_for_selection()

    def _refresh_data(self, doc=None, doc_switched=False):
        if doc is None:
            doc = self._get_active_doc()
        current_doc_identity = self._doc_identity(doc)
        try:
            _path, data = load_active_yaml_data(doc)
        except Exception as exc:
            # Keep cached data only when the same document is still active.
            same_doc_cache = bool(self._raw_data and self._repo and self._data_doc_identity == current_doc_identity)
            if same_doc_cache and not doc_switched:
                self._set_status("Using cached YAML (load failed: {}).".format(exc))
                try:
                    self._refresh_profile_choices()
                except Exception:
                    pass
                return
            self._raw_data = {}
            self._repo = None
            self._choice_map = {}
            self._profile_type_assets = {}
            self._data_doc_identity = None
            self._set_status("Failed to load active YAML: {}".format(exc))
            if self._profile_combo is not None:
                self._profile_combo.ItemsSource = []
            self._update_profile_type_list([])
            return

        self._raw_data = data
        self._data_doc_identity = current_doc_identity
        cleaned_profiles = _cleaned_profiles_from_raw(data)
        self._repo = _build_repository_from_profiles(cleaned_profiles)
        self._profile_type_assets = _build_profile_type_assets(cleaned_profiles)
        self._symbol_cache = {}
        self._symbol_lookup = None
        self._symbol_lookup_info = {}
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
        self._update_profile_type_list_for_selection()

    def _update_profile_type_list_for_selection(self):
        cad_choice = self._selected_profile()
        labels = []
        if cad_choice and self._repo:
            try:
                labels = list(self._repo.labels_for_cad(cad_choice) or [])
            except Exception:
                labels = []
        self._update_profile_type_list(labels, cad_choice=cad_choice)

    def _selected_profile(self):
        if self._profile_combo is None:
            return None
        label = self._profile_combo.SelectedItem
        if label is None:
            label = getattr(self._profile_combo, "Text", None)
        if not label:
            return None
        return self._choice_map.get(label, label)

    def _update_profile_type_list(self, labels, cad_choice=None):
        if self._profile_types_list is None:
            return
        if not labels:
            try:
                self._profile_types_list.ItemsSource = []
            except Exception:
                pass
            return
        assets_by_label = {}
        if cad_choice:
            assets_by_label = self._profile_type_assets.get(cad_choice) or {}
        items = []
        for label in labels:
            if not label:
                continue
            items.append(_format_profile_type_item(label, assets_by_label.get(label)))
        try:
            self._profile_types_list.ItemsSource = items
        except Exception:
            pass

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

    def _on_place_on_reference(self, sender, args):
        if not self._repo:
            self._set_status("No active YAML loaded.")
            return
        cad_choice = self._selected_profile()
        if not cad_choice:
            self._set_status("Select a profile to place.")
            return
        try:
            forms.alert(
                "Select reference elements in the model, then press Enter or Finish.\n"
                "The profile will be placed on each selected element.",
                title="Place on Reference Elements",
                warn_icon=False,
            )
        except Exception:
            pass
        self._place_on_reference_handler.set_payload(self._raw_data, cad_choice)
        self._place_on_reference_event.Raise()
        self._set_status("Placement requested for '{}' on reference elements.".format(cad_choice))


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
        panel._sync_active_document(force=True)

