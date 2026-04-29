# -*- coding: utf-8 -*-
"""
Synced-relationship audit.

For every placed child element with an ``Element_Linker`` payload, we
walk the LED's ``parameters`` dict; any ``parent_parameter`` or
``sibling_parameter`` directive defines an *expected* value. If the
child's actual value diverges, we record a Conflict. The rewritten UI
groups conflicts as ``Profile -> LED -> Conflict`` so users can see the
shape of their constraints.

The detector is decoupled from any UI: ``detect_conflicts`` returns
plain dicts, and ``apply_resolution`` does the parameter write.
"""

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    FamilyInstance,
    FilteredElementCollector,
    Group,
    StorageType,
)

import directives as _dir
import element_linker_io as _el_io
import profile_model


# ---------------------------------------------------------------------
# Conflict / Decision data classes
# ---------------------------------------------------------------------

class Conflict(object):
    """One detected mismatch."""

    UPDATE_CHILD = "update_child"
    UPDATE_PARENT = "update_parent"
    SKIP = "skip"

    def __init__(self, profile_id, profile_name, led_id, led_label,
                 element_id, parameter_name, kind, expected_value,
                 actual_value, target_param_name, target_element_id):
        self.profile_id = profile_id
        self.profile_name = profile_name
        self.led_id = led_id
        self.led_label = led_label
        self.element_id = element_id
        self.parameter_name = parameter_name
        self.kind = kind  # "parent" | "sibling"
        self.expected_value = expected_value
        self.actual_value = actual_value
        self.target_param_name = target_param_name
        self.target_element_id = target_element_id

    @property
    def key(self):
        return (self.profile_id, self.led_id, self.element_id, self.parameter_name)

    def to_display_dict(self):
        return {
            "profile": "{} ({})".format(self.profile_name, self.profile_id),
            "led": "{} ({})".format(self.led_label, self.led_id),
            "element_id": self.element_id,
            "parameter": self.parameter_name,
            "kind": self.kind,
            "expected": self.expected_value,
            "actual": self.actual_value,
        }


# ---------------------------------------------------------------------
# Parameter read / write
# ---------------------------------------------------------------------

def _read_param_value(elem, name):
    if elem is None:
        return None
    param = elem.LookupParameter(name)
    if param is None:
        return None
    storage = param.StorageType
    if storage == StorageType.String:
        return param.AsString()
    if storage == StorageType.Integer:
        return param.AsInteger()
    if storage == StorageType.Double:
        return param.AsDouble()
    if storage == StorageType.ElementId:
        eid = param.AsElementId()
        return getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)
    return param.AsValueString() or param.AsString()


def _write_param_value(elem, name, value):
    """Write a value, coerced to the parameter's storage type. Returns True on success."""
    if elem is None:
        return False
    param = elem.LookupParameter(name)
    if param is None or param.IsReadOnly:
        return False
    storage = param.StorageType
    try:
        if storage == StorageType.String:
            param.Set("" if value is None else str(value))
        elif storage == StorageType.Integer:
            param.Set(int(value) if value is not None else 0)
        elif storage == StorageType.Double:
            param.Set(float(value) if value is not None else 0.0)
        else:
            return False
    except Exception:
        return False
    return True


# ---------------------------------------------------------------------
# Detection
# ---------------------------------------------------------------------

def _values_match(a, b):
    if a is None and b is None:
        return True
    if isinstance(a, float) or isinstance(b, float):
        try:
            return abs(float(a) - float(b)) < 1e-6
        except (TypeError, ValueError):
            pass
    return str(a or "").strip().lower() == str(b or "").strip().lower()


def _collect_placed_elements_with_linker(doc):
    """Return ``{(profile_id_or_None, set_id, led_id): [(elem, linker), ...]}``."""
    out = {}
    for klass in (FamilyInstance, Group):
        collector = FilteredElementCollector(doc).OfClass(klass).WhereElementIsNotElementType()
        for elem in collector:
            linker = _el_io.read_from_element(elem)
            if linker is None or not linker.set_id or not linker.led_id:
                continue
            key = (linker.set_id, linker.led_id)
            out.setdefault(key, []).append((elem, linker))
    return out


def _build_sibling_lookup(siblings_in_set):
    """Given ``[(elem, linker), ...]`` for one set, build a callable
    ``(led_id, param_name) -> value | None``."""
    by_led = {}
    for elem, linker in siblings_in_set:
        by_led.setdefault(linker.led_id, []).append(elem)

    def lookup(led_id, param_name):
        for elem in by_led.get(led_id, []):
            v = _read_param_value(elem, param_name)
            if v is not None:
                return v
        return None
    return lookup


def detect_conflicts(doc, profile_data):
    """Return a list of ``Conflict``."""
    conflicts = []
    pdoc = profile_model.ProfileDocument(profile_data)

    # Group all placed elements by set so we can resolve sibling references.
    placed = _collect_placed_elements_with_linker(doc)
    siblings_in_set = {}
    for (set_id, led_id), entries in placed.items():
        siblings_in_set.setdefault(set_id, []).extend(entries)

    for profile in pdoc.profiles:
        for linked_set in profile.linked_sets:
            sibling_lookup = _build_sibling_lookup(
                siblings_in_set.get(linked_set.id, [])
            )
            for led in linked_set.leds:
                params = led.parameters or {}
                placed_entries = placed.get((linked_set.id, led.id), [])
                if not placed_entries:
                    continue
                # Each placed instance gets its own audit pass — the
                # parent for resolution is whichever element the linker
                # points at as the parent.
                for elem, linker in placed_entries:
                    parent_elem = None
                    if linker.parent_element_id:
                        parent_elem = doc.GetElement(_make_element_id(doc, linker.parent_element_id))
                    parent_lookup = _build_parent_lookup(parent_elem)
                    for param_name, value in params.items():
                        kind = _dir.directive_kind(value)
                        if kind == "static":
                            continue
                        found, expected = _dir.resolve_expected_value(
                            value, parent_lookup, sibling_lookup
                        )
                        if not found:
                            continue
                        actual = _read_param_value(elem, param_name)
                        if _values_match(actual, expected):
                            continue
                        conflicts.append(Conflict(
                            profile_id=profile.id,
                            profile_name=profile.name,
                            led_id=led.id,
                            led_label=led.label,
                            element_id=_id_value(elem),
                            parameter_name=param_name,
                            kind=kind,
                            expected_value=expected,
                            actual_value=actual,
                            target_param_name=_target_param_name(value, kind),
                            target_element_id=(
                                _id_value(parent_elem) if kind == "parent"
                                else _sibling_target_id(value, siblings_in_set.get(linked_set.id, []))
                            ),
                        ))
    return conflicts


def _make_element_id(doc, value):
    from Autodesk.Revit.DB import ElementId
    try:
        return ElementId(int(value))
    except Exception:
        return ElementId.InvalidElementId


def _id_value(elem):
    if elem is None:
        return None
    eid = elem.Id
    return getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)


def _build_parent_lookup(parent_elem):
    def lookup(param_name):
        return _read_param_value(parent_elem, param_name)
    return lookup


def _target_param_name(directive_value, kind):
    if kind == "parent":
        return _dir.parent_param_name(directive_value)
    if kind == "sibling":
        target = _dir.sibling_target(directive_value)
        return target[1] if target else None
    return None


def _sibling_target_id(directive_value, set_entries):
    target = _dir.sibling_target(directive_value)
    if target is None:
        return None
    led_id, _ = target
    for elem, linker in set_entries:
        if linker.led_id == led_id:
            return _id_value(elem)
    return None


# ---------------------------------------------------------------------
# Resolution
# ---------------------------------------------------------------------

def apply_resolution(doc, conflict, action):
    """Mutate the model to apply ``action`` to ``conflict``.

    Caller manages the transaction.
    """
    if action == Conflict.SKIP or action is None:
        return False
    if action == Conflict.UPDATE_CHILD:
        elem = doc.GetElement(_make_element_id(doc, conflict.element_id))
        return _write_param_value(elem, conflict.parameter_name, conflict.expected_value)
    if action == Conflict.UPDATE_PARENT:
        if conflict.target_element_id is None or conflict.target_param_name is None:
            return False
        target = doc.GetElement(_make_element_id(doc, conflict.target_element_id))
        return _write_param_value(target, conflict.target_param_name, conflict.actual_value)
    return False
