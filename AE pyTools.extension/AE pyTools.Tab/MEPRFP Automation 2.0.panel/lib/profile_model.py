# -*- coding: utf-8 -*-
"""
Typed wrapper classes for equipment-definition YAML.

Each class wraps a dict (the underlying YAML node) and exposes its
fields as Python attributes. Wrappers *own* their dict — mutating the
wrapper mutates the source. ``to_dict()`` returns the underlying dict
verbatim so the YAML round-trip stays lossless.

Field set is derived from the v4 schema observed in
HEB_profiles_V4_MODIFIED_*.yaml plus the constants in profile_schema.py.
Unknown fields are preserved through ``_extra`` (the underlying dict),
so adding new YAML fields doesn't require updating these classes.
"""


# ---------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------

def _str_or_none(value):
    if value is None:
        return None
    return str(value)


def _bool_or(value, default):
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered == "true":
            return True
        if lowered == "false":
            return False
    return bool(value)


def _ensure_list(value):
    if value is None:
        return []
    if isinstance(value, list):
        return value
    return [value]


# ---------------------------------------------------------------------
# Wrappers
# ---------------------------------------------------------------------

class _DictBacked(object):
    """Base class: wrapper owns ``data`` (the underlying dict)."""

    def __init__(self, data=None):
        self._data = data if data is not None else {}

    def to_dict(self):
        return self._data

    def __eq__(self, other):
        if not isinstance(other, _DictBacked):
            return NotImplemented
        return self._data == other._data

    def __ne__(self, other):
        result = self.__eq__(other)
        return result if result is NotImplemented else not result


class Offset(_DictBacked):
    """An ``{x_inches, y_inches, z_inches, rotation_deg}`` triple."""

    @property
    def x_inches(self):
        return float(self._data.get("x_inches") or 0.0)

    @x_inches.setter
    def x_inches(self, value):
        self._data["x_inches"] = float(value)

    @property
    def y_inches(self):
        return float(self._data.get("y_inches") or 0.0)

    @y_inches.setter
    def y_inches(self, value):
        self._data["y_inches"] = float(value)

    @property
    def z_inches(self):
        return float(self._data.get("z_inches") or 0.0)

    @z_inches.setter
    def z_inches(self, value):
        self._data["z_inches"] = float(value)

    @property
    def rotation_deg(self):
        return float(self._data.get("rotation_deg") or 0.0)

    @rotation_deg.setter
    def rotation_deg(self, value):
        self._data["rotation_deg"] = float(value)


class Tag(_DictBacked):
    """An entry from ``led.tags``. Tag offsets are a single dict, not a list."""

    @property
    def category_name(self):
        return _str_or_none(self._data.get("category_name"))

    @property
    def family_name(self):
        return _str_or_none(self._data.get("family_name"))

    @property
    def type_name(self):
        return _str_or_none(self._data.get("type_name"))

    @property
    def parameters(self):
        return self._data.setdefault("parameters", {})

    @property
    def offset(self):
        d = self._data.setdefault("offsets", {})
        return Offset(d) if isinstance(d, dict) else None


class Annotation(_DictBacked):
    """A single ``annotations[*]`` entry under a LED.

    ``kind``    "tag" / "keynote" / "text_note"
    ``label``   "Family : Type" or other display string
    ``parameters``  {name: value} dict
    ``offset``  ``Offset`` object — relative to the host fixture, NOT the
                profile parent
    """

    @property
    def id(self):
        return _str_or_none(self._data.get("id"))

    @property
    def kind(self):
        return _str_or_none(self._data.get("kind"))

    @property
    def label(self):
        return _str_or_none(self._data.get("label"))

    @property
    def category_name(self):
        return _str_or_none(self._data.get("category_name"))

    @property
    def family_name(self):
        return _str_or_none(self._data.get("family_name"))

    @property
    def type_name(self):
        return _str_or_none(self._data.get("type_name"))

    @property
    def parameters(self):
        return self._data.setdefault("parameters", {})

    @property
    def offset(self):
        d = self._data.setdefault("offsets", {})
        if not isinstance(d, dict):
            d = {}
            self._data["offsets"] = d
        return Offset(d)


class LED(_DictBacked):
    """A linked-element-definition entry."""

    @property
    def id(self):
        return _str_or_none(self._data.get("id"))

    @property
    def label(self):
        return _str_or_none(self._data.get("label"))

    @property
    def category(self):
        return _str_or_none(self._data.get("category"))

    @property
    def is_group(self):
        return _bool_or(self._data.get("is_group"), False)

    @property
    def is_parent_anchor(self):
        return _bool_or(self._data.get("is_parent_anchor"), False)

    @property
    def parameters(self):
        return self._data.setdefault("parameters", {})

    @property
    def offsets(self):
        """LED offsets are a *list* of Offset entries."""
        raw = self._data.setdefault("offsets", [])
        return [Offset(o) for o in _ensure_list(raw)]

    @property
    def annotations(self):
        """Unified ``annotations[*]`` list. Legacy LEDs that only have
        the older ``tags``/``keynotes``/``text_notes`` peer lists get
        synthesised on the fly so callers see one consistent shape."""
        raw = self._data.get("annotations")
        if isinstance(raw, list):
            return [Annotation(a) for a in raw]
        # Legacy fallback: synthesize from peer lists.
        synth = []
        for kind, key in (("tag", "tags"), ("keynote", "keynotes"),
                          ("text_note", "text_notes")):
            for entry in (self._data.get(key) or []):
                if not isinstance(entry, dict):
                    continue
                copy = dict(entry)
                copy.setdefault("kind", kind)
                synth.append(copy)
        return [Annotation(a) for a in synth]

    @property
    def tags(self):
        raw = self._data.setdefault("tags", [])
        return [Tag(t) for t in _ensure_list(raw)]

    @property
    def keynotes(self):
        raw = self._data.setdefault("keynotes", [])
        return [Tag(k) for k in _ensure_list(raw)]

    @property
    def text_notes(self):
        raw = self._data.setdefault("text_notes", [])
        return [Tag(n) for n in _ensure_list(raw)]


class LinkedSet(_DictBacked):
    @property
    def id(self):
        return _str_or_none(self._data.get("id"))

    @property
    def name(self):
        return _str_or_none(self._data.get("name"))

    @property
    def leds(self):
        raw = self._data.setdefault("linked_element_definitions", [])
        return [LED(l) for l in _ensure_list(raw)]


class ParentFilter(_DictBacked):
    @property
    def category(self):
        return _str_or_none(self._data.get("category"))

    @property
    def family_name_pattern(self):
        return _str_or_none(self._data.get("family_name_pattern"))

    @property
    def type_name_pattern(self):
        return _str_or_none(self._data.get("type_name_pattern"))

    @property
    def parameter_filters(self):
        return self._data.setdefault("parameter_filters", {})


class Profile(_DictBacked):
    """An ``equipment_definitions[*]`` entry."""

    @property
    def id(self):
        return _str_or_none(self._data.get("id"))

    @property
    def name(self):
        return _str_or_none(self._data.get("name"))

    @property
    def schema_version(self):
        v = self._data.get("schema_version")
        if v is None:
            return None
        try:
            return int(v)
        except (ValueError, TypeError):
            return None

    @property
    def parent_filter(self):
        d = self._data.setdefault("parent_filter", {})
        return ParentFilter(d)

    @property
    def linked_sets(self):
        raw = self._data.setdefault("linked_sets", [])
        return [LinkedSet(s) for s in _ensure_list(raw)]

    @property
    def equipment_properties(self):
        return self._data.setdefault("equipment_properties", {})

    @property
    def allow_parentless(self):
        return _bool_or(self._data.get("allow_parentless"), True)

    @property
    def allow_unmatched_parents(self):
        return _bool_or(self._data.get("allow_unmatched_parents"), True)

    @property
    def prompt_on_parent_mismatch(self):
        return _bool_or(self._data.get("prompt_on_parent_mismatch"), False)

    @property
    def truth_source_id(self):
        return _str_or_none(self._data.get("ced_truth_source_id"))

    @property
    def truth_source_name(self):
        return _str_or_none(self._data.get("ced_truth_source_name"))


class ProfileDocument(_DictBacked):
    """Root of the YAML document."""

    @property
    def schema_version(self):
        v = self._data.get("schema_version")
        if v is None:
            return None
        try:
            return int(v)
        except (ValueError, TypeError):
            return None

    @property
    def profiles(self):
        raw = self._data.setdefault("equipment_definitions", [])
        return [Profile(p) for p in _ensure_list(raw)]

    def find_profile_by_id(self, profile_id):
        for p in self.profiles:
            if p.id == profile_id:
                return p
        return None

    def find_profile_by_name(self, name):
        for p in self.profiles:
            if p.name == name:
                return p
        return None
