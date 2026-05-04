# -*- coding: utf-8 -*-
"""
Orchestration layer for SuperCircuit (MEPRFP Automation 2.0).

Pipeline overview:

    1. ``collect_items(doc, client, scope)``
        Walk every electrical-fixture / data-device / mech-control
        FamilyInstance in the doc (or the user's selection), read
        Element_Linker first (CKT_Panel / CKT_Circuit / load) then
        fall back to live Revit parameters; emit one ``CircuitItem``
        per element.

    2. ``client.enrich_item(item, context)`` for each item.
        Default no-op. HEB hooks panel-distance ranking, space-name
        load decoration, casecontroller bucketing.

    3. ``classify_circuit_token`` populates ``item.bucket`` and
       ``item.circuit_token``. Position-rule matches and
       ``post_enrich_classify`` overrides take precedence.

    4. ``circuit_grouping.assemble_groups(items)`` produces the
       ``CircuitGroup`` plan list the UI shows.

    5. The UI lets the user edit per-row panel / circuit / load.
       On Apply it calls ``CircuitApplyGateway.request_apply`` which
       hops to the Revit thread and creates the systems.

Pure orchestration — no WPF imports, only Revit API + the lib's own
modules. The UI lives in ``circuit_window.py``.
"""

import math

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    LocationPoint,
)

import circuit_apply as _apply
import circuit_clients as _clients
import circuit_grouping as _grouping
import element_linker_io as _el_io


# ---------------------------------------------------------------------
# Scope
# ---------------------------------------------------------------------

SCOPE_ALL = "all"
SCOPE_SELECTION = "selection"


# Categories the workflow considers "electrical" for circuiting.
_TARGET_CATEGORIES = (
    BuiltInCategory.OST_ElectricalFixtures,
    BuiltInCategory.OST_DataDevices,
    BuiltInCategory.OST_CommunicationDevices,
    BuiltInCategory.OST_FireAlarmDevices,
    BuiltInCategory.OST_NurseCallDevices,
    BuiltInCategory.OST_SecurityDevices,
    BuiltInCategory.OST_TelephoneDevices,
    BuiltInCategory.OST_MechanicalControlDevices,
)

# Categories explicitly excluded (legacy V5 dropped lighting on
# purpose — those are circuited by a separate workflow).
_EXCLUDED_CATEGORIES = frozenset({
    int(BuiltInCategory.OST_LightingDevices),
    int(BuiltInCategory.OST_LightingFixtures),
})


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------

def _id_value(elem_or_id):
    if elem_or_id is None:
        return None
    eid = getattr(elem_or_id, "Id", None) or elem_or_id
    return getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)


def _category_id(elem):
    cat = getattr(elem, "Category", None)
    if cat is None:
        return None
    return _id_value(cat)


def _get_param_string(elem, name):
    if elem is None:
        return ""
    try:
        p = elem.LookupParameter(name)
    except Exception:
        return ""
    if p is None:
        return ""
    try:
        if not p.HasValue:
            return ""
    except Exception:
        pass
    try:
        s = p.AsString()
        if s is not None:
            return s
    except Exception:
        pass
    try:
        return p.AsValueString() or ""
    except Exception:
        return ""


def _get_param_int(elem, name):
    if elem is None:
        return None
    try:
        p = elem.LookupParameter(name)
    except Exception:
        return None
    if p is None:
        return None
    try:
        v = p.AsInteger()
        return int(v)
    except Exception:
        try:
            s = p.AsString() or ""
            return int(float(s))
        except Exception:
            return None


def _get_param_float(elem, name):
    if elem is None:
        return None
    try:
        p = elem.LookupParameter(name)
    except Exception:
        return None
    if p is None:
        return None
    try:
        v = p.AsDouble()
        return float(v)
    except Exception:
        try:
            s = p.AsString() or ""
            return float(s)
        except Exception:
            return None


def _location_xyz_tuple(elem):
    loc = getattr(elem, "Location", None)
    if not isinstance(loc, LocationPoint):
        return None
    try:
        p = loc.Point
        if p is None:
            return None
        return (p.X, p.Y, p.Z)
    except Exception:
        return None


def _family_and_type(elem):
    if not isinstance(elem, FamilyInstance):
        return ("", "")
    sym = getattr(elem, "Symbol", None)
    if sym is None:
        return ("", "")
    family = getattr(sym, "Family", None)
    fam = getattr(family, "Name", "") if family is not None else ""
    typ = getattr(sym, "Name", "") or ""
    return (fam, typ)


def _read_linker_or_none(elem):
    try:
        return _el_io.read_from_element(elem)
    except Exception:
        return None


# ---------------------------------------------------------------------
# YAML / profile lookup
# ---------------------------------------------------------------------

def _build_led_index(profile_data):
    """{led_id: (profile, set_dict, led)}."""
    out = {}
    if not isinstance(profile_data, dict):
        return out
    for profile in profile_data.get("equipment_definitions") or []:
        if not isinstance(profile, dict):
            continue
        for set_dict in profile.get("linked_sets") or []:
            if not isinstance(set_dict, dict):
                continue
            for led in set_dict.get("linked_element_definitions") or []:
                if isinstance(led, dict) and led.get("id"):
                    out[led["id"]] = (profile, set_dict, led)
    return out


# ---------------------------------------------------------------------
# Item collection
# ---------------------------------------------------------------------

def collect_items(doc, client, profile_data=None,
                  scope=SCOPE_ALL, selected_element_ids=None):
    """Walk the doc (or selection) and produce ``CircuitItem`` records.

    Element_Linker is consulted first; live Revit parameters fill in
    anything the linker doesn't carry. The client's ``should_circuit``
    hook is the final filter — it can drop e.g. specific family types
    a client isn't supposed to circuit.
    """
    led_index = _build_led_index(profile_data) if profile_data else {}
    out = []
    elements = _enumerate_target_elements(doc, scope, selected_element_ids)
    for elem in elements:
        item = _build_item(elem, client, led_index)
        if item is None:
            continue
        if not client.should_circuit(elem, item):
            continue
        out.append(item)
    return out


def _enumerate_target_elements(doc, scope, selected_element_ids):
    if scope == SCOPE_SELECTION and selected_element_ids:
        for eid in selected_element_ids:
            try:
                elem = doc.GetElement(eid)
            except Exception:
                continue
            if elem is None:
                continue
            cat_id = _category_id(elem)
            if cat_id is None or cat_id in _EXCLUDED_CATEGORIES:
                continue
            if isinstance(elem, FamilyInstance):
                yield elem
        return

    for cat in _TARGET_CATEGORIES:
        try:
            collector = (
                FilteredElementCollector(doc)
                .OfCategory(cat)
                .OfClass(FamilyInstance)
                .WhereElementIsNotElementType()
            )
        except Exception:
            continue
        for elem in collector:
            cat_id = _category_id(elem)
            if cat_id is not None and cat_id in _EXCLUDED_CATEGORIES:
                continue
            yield elem


def _build_item(elem, client, led_index):
    """Construct one ``CircuitItem`` from a placed element."""
    if elem is None:
        return None

    fam, typ = _family_and_type(elem)
    linker = _read_linker_or_none(elem)
    led, profile = None, None
    if linker is not None and linker.led_id:
        entry = led_index.get(linker.led_id)
        if entry is not None:
            profile, _set_dict, led = entry

    # Source-of-truth resolution:
    #   1. Element_Linker payload (preferred — placement engine writes it).
    #   2. Live Revit parameter on the placed instance.
    #   3. YAML LED.parameters as last-chance fallback.

    panel_raw = ""
    circuit_raw = ""
    if linker is not None:
        if linker.ckt_panel:
            panel_raw = str(linker.ckt_panel)
        if linker.ckt_circuit_number:
            circuit_raw = str(linker.ckt_circuit_number)
    if not panel_raw:
        panel_raw = _get_param_string(elem, "CKT_Panel_CEDT")
    if not circuit_raw:
        circuit_raw = _get_param_string(elem, "CKT_Circuit Number_CEDT")
    load_raw = _get_param_string(elem, "CKT_Load Name_CEDT")

    if (not panel_raw or not circuit_raw) and isinstance(led, dict):
        led_params = led.get("parameters") or {}
        if isinstance(led_params, dict):
            if not panel_raw:
                panel_raw = str(led_params.get("CKT_Panel_CEDT") or "")
            if not circuit_raw:
                circuit_raw = str(led_params.get("CKT_Circuit Number_CEDT") or "")
            if not load_raw:
                load_raw = str(led_params.get("CKT_Load Name_CEDT") or "")

    panel_choices = client.parse_panel_string(panel_raw)
    panel_name = panel_choices[0] if panel_choices else ""
    bucket, circuit_token = client.classify_circuit_token(circuit_raw)

    rating = _get_param_float(elem, "CKT_Rating_CED")
    if rating is None:
        # Some templates store rating as a string in display units
        # (e.g. ``"20 A"``); fall back to the parameter's value-string
        # path so we still capture the number.
        rating_text = _get_param_string(elem, "CKT_Rating_CED")
        if rating_text:
            try:
                rating = float(rating_text.split()[0])
            except (ValueError, IndexError):
                rating = None
    voltage = _get_param_float(elem, "Voltage_CED")
    poles = _get_param_int(elem, "Number of Poles_CED") or 1
    schedule_notes = _get_param_string(elem, "CKT_Schedule Notes_CEDT")

    # YAML LED parameters as last-chance fallbacks for rating / notes.
    if isinstance(led, dict):
        led_params = led.get("parameters") or {}
        if isinstance(led_params, dict):
            if rating is None:
                yaml_rating = led_params.get("CKT_Rating_CED")
                if yaml_rating not in (None, ""):
                    try:
                        rating = float(yaml_rating)
                    except (TypeError, ValueError):
                        try:
                            rating = float(str(yaml_rating).split()[0])
                        except (ValueError, IndexError):
                            rating = None
            if not schedule_notes:
                yaml_notes = led_params.get("CKT_Schedule Notes_CEDT")
                if yaml_notes:
                    schedule_notes = str(yaml_notes)

    item = _grouping.CircuitItem(
        element=elem,
        element_id=_id_value(elem),
        family_name=fam,
        type_name=typ,
        linker=linker,
        led=led,
        profile=profile,
        panel_raw=panel_raw,
        circuit_number_raw=circuit_raw,
        load_name_raw=load_raw,
        panel_choices=panel_choices,
        panel_name=panel_name,
        circuit_token=circuit_token,
        load_name=load_raw,
        rating=rating,
        voltage=voltage,
        poles=poles,
        schedule_notes=schedule_notes,
        world_pt=_location_xyz_tuple(elem),
        parent_element_id=(linker.parent_element_id if linker is not None else None),
        bucket=bucket,
    )
    return item


# ---------------------------------------------------------------------
# Enrichment
# ---------------------------------------------------------------------

def enrich_items(doc, items, client, panel_index=None):
    """Apply the client's enrichment + position-rule + post-classify
    passes. Mutates items in place.

    Order:

        1. ``client.enrich_item`` (per-item, free-form).
        2. ``client.rank_panel_choices`` (sorts ``panel_choices``;
           default no-op).
        3. ``client.decorate_load_name`` (writes ``item.load_name``).
        4. Position-rule check — if a rule matches the raw circuit
           number, the rule's ``label`` (when set) overrides the load
           name and the item lands in a custom ``position_<keyword>``
           bucket so the grouper batches by spatial proximity within
           the rule.
        5. ``client.post_enrich_classify`` — final override (e.g. HEB
           ``CASECONTROLLER_<prefix>``).
    """
    if not items:
        return items
    if panel_index is None:
        panel_index = _apply.collect_panel_index(doc)

    # Build a panel "world point" lookup for distance-based ranking.
    panels_by_name = {}
    for key, panel_elem in (panel_index or {}).items():
        pt = _apply.panel_world_pt(panel_elem)
        # Wrap into a thin object exposing world_pt for the client hook.
        class _PanelView(object):
            __slots__ = ("world_pt", "element")
        view = _PanelView()
        view.world_pt = pt
        view.element = panel_elem
        panels_by_name[key] = view

    context = {
        "panels_by_name": panels_by_name,
        "space_name_for_item": lambda _item: None,  # populated below if needed
    }

    for item in items:
        client.enrich_item(item, context)
        ranked = client.rank_panel_choices(item, context)
        if ranked is not None:
            item.panel_choices = list(ranked)
            if item.panel_choices:
                item.panel_name = item.panel_choices[0]

        # position-rule pass
        rule = client.match_position_rule(item.circuit_number_raw)
        if rule is not None:
            keyword = (rule.get("keyword") or "").strip().upper()
            if keyword:
                item.bucket = "position_{}".format(keyword.lower().replace(" ", "_"))
                item.circuit_token = keyword
                if rule.get("label"):
                    item.load_name = str(rule["label"])

        decorated = client.decorate_load_name(item, context)
        if decorated is not None:
            item.load_name = decorated

        override = client.post_enrich_classify(item)
        if override:
            new_bucket, new_token = override
            if new_bucket:
                item.bucket = new_bucket
            if new_token:
                item.circuit_token = new_token

    return items


# ---------------------------------------------------------------------
# Top-level orchestration
# ---------------------------------------------------------------------

class CircuitRun(object):
    """Carries the state of one SuperCircuit invocation between the
    workflow, the preview UI, and the apply gateway."""

    __slots__ = (
        "doc", "client", "scope", "selected_element_ids",
        "profile_data", "panel_index",
        "items", "groups",
        "gateway",
    )

    def __init__(self, doc, client, scope=SCOPE_ALL,
                 selected_element_ids=None, profile_data=None):
        self.doc = doc
        self.client = client
        self.scope = scope
        self.selected_element_ids = list(selected_element_ids or [])
        self.profile_data = profile_data
        self.panel_index = {}
        self.items = []
        self.groups = []
        # Resolve the per-Revit-session ExternalEvent gateway *now*,
        # while the pushbutton script is still inside the Revit API
        # execution context. ``ExternalEvent.Create`` is only valid in
        # that context; creating it lazily from a later WPF event
        # handler raises "Attempting to create an ExternalEvent
        # outside of a standard API execution". The gateway is
        # cached at module level so subsequent SuperCircuit runs
        # reuse it instead of registering a duplicate handler type.
        try:
            self.gateway = _apply.get_or_create_gateway()
        except Exception:
            self.gateway = None

    # ----- pipeline -------------------------------------------------

    def collect(self):
        """Run collect + enrich passes. Populates ``items`` and
        ``panel_index``. Doesn't touch ``groups``."""
        self.panel_index = _apply.collect_panel_index(self.doc)
        self.items = collect_items(
            self.doc, self.client,
            profile_data=self.profile_data,
            scope=self.scope,
            selected_element_ids=self.selected_element_ids,
        )
        enrich_items(self.doc, self.items, self.client, self.panel_index)
        return self.items

    def assemble(self):
        """Build groups from the current items. Re-runnable after the
        UI mutates ``user_panel`` / ``user_circuit_token`` /
        ``user_load_name`` on items."""
        self.groups = _grouping.assemble_groups(self.items)
        # Resolve panel_element + sort by client's load priority.
        for group in self.groups:
            key = (group.panel_name or "").strip().lower()
            group.panel_element = self.panel_index.get(key)
        self.groups.sort(key=lambda g: (
            self.client.get_load_priority(g.load_name),
            (g.panel_name or "").lower(),
            (g.circuit_token or "").lower(),
        ))
        return self.groups

    # ----- apply ----------------------------------------------------

    def apply_async(self, on_complete=None, groups=None):
        """Hop to the Revit thread via ``ExternalEvent`` and create
        circuits for every non-review group. ``on_complete`` is called
        with a ``CircuitApplyResult`` after the transaction commits.

        ``groups`` overrides the default of "every group on the run" —
        the UI passes the keyword-filtered subset so users can scope a
        creation pass to e.g. just DEDICATED rows.
        """
        if self.gateway is None:
            raise RuntimeError(
                "ExternalEvent gateway was never created. CircuitRun must "
                "be constructed inside the pushbutton script's API "
                "context — re-launch the SuperCircuit V5 button."
            )
        # Re-collect the panel index right before apply so panels added
        # since the run started are visible.
        if not self.panel_index:
            self.panel_index = _apply.collect_panel_index(self.doc)
        target_groups = list(groups) if groups is not None else list(self.groups)
        # Resolve panel_element on every group with current index.
        for g in target_groups:
            key = (g.panel_name or "").strip().lower()
            g.panel_element = self.panel_index.get(key)
        self.gateway.request_apply(
            self.doc, target_groups,
            transaction_name="SuperCircuit (MEPRFP 2.0)",
            on_complete=on_complete,
        )


# ---------------------------------------------------------------------
# Run keyword filter (used by the UI dropdown)
# ---------------------------------------------------------------------

def filter_groups_by_keyword(groups, keyword, client):
    """Return groups whose tokens include ``keyword``.

    ``token`` membership is checked against the upper-cased
    ``circuit_token`` and the bucket label, plus a few synonyms for
    BYPARENT-family keys so the filter dropdown matches the legacy V5
    behaviour.
    """
    if not keyword or not groups:
        return list(groups or [])
    target = keyword.strip().upper()
    if not target:
        return list(groups)
    out = []
    for group in groups:
        token = (group.circuit_token or "").strip().upper()
        bucket = (group.bucket or "").strip().lower()
        synonyms = []
        if bucket == _grouping.BUCKET_DEDICATED:
            synonyms.append("DEDICATED")
        elif bucket == _grouping.BUCKET_BYPARENT:
            synonyms.append("BYPARENT")
        elif bucket == _grouping.BUCKET_SECONDBYPARENT:
            synonyms.append("SECONDBYPARENT")
        if target == token or target in synonyms:
            out.append(group)
    return out
