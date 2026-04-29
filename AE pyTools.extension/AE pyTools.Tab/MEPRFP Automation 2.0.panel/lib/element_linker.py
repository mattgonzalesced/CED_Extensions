# -*- coding: utf-8 -*-
"""
Element_Linker payload codec.

The Element_Linker parameter on a placed instance carries metadata that
points the element back to its YAML profile origin. MEPRFP 2.0 stores
this as a JSON object; the legacy bespoke "Key: value" text format is
read by ``from_legacy_text`` so existing project data can be migrated.

JSON schema (codec version 1)::

    {
      "v": 1,
      "led_id": str | null,
      "set_id": str | null,
      "location_ft": [x, y, z] | null,
      "rotation_deg": float | null,
      "parent_rotation_deg": float | null,
      "parent_element_id": int | null,
      "level_id": int | null,
      "element_id": int | null,
      "facing": [x, y, z] | null,
      "host_name": str | null,
      "parent_location_ft": [x, y, z] | null
    }

The Revit parameter name is fixed at ``"Element_Linker"``. The legacy
codebase also honoured ``"Element_Linker Parameter"``; 2.0 only writes
the canonical name. Migration tooling is responsible for reading either.
"""

import json
import re


PARAMETER_NAME = "Element_Linker"
CODEC_VERSION = 1


_FIELDS = (
    "led_id",
    "set_id",
    "location_ft",
    "rotation_deg",
    "parent_rotation_deg",
    "parent_element_id",
    "level_id",
    "element_id",
    "facing",
    "host_name",
    "parent_location_ft",
)


_LEGACY_FIELD_MAP = {
    "Linked Element Definition ID": "led_id",
    "Set Definition ID": "set_id",
    "Location XYZ (ft)": "location_ft",
    "Rotation (deg)": "rotation_deg",
    "Parent Rotation (deg)": "parent_rotation_deg",
    "Parent ElementId": "parent_element_id",
    "Parent Element ID": "parent_element_id",
    "LevelId": "level_id",
    "Level Id": "level_id",
    "ElementId": "element_id",
    "Element ID": "element_id",
    "Element Id": "element_id",
    "FacingOrientation": "facing",
    "Host Name": "host_name",
    "Parent_location": "parent_location_ft",
}

_TUPLE_FIELDS = {"location_ft", "facing", "parent_location_ft"}
_INT_FIELDS = {"parent_element_id", "level_id", "element_id"}
_FLOAT_FIELDS = {"rotation_deg", "parent_rotation_deg"}


# Built once: a regex that matches any known legacy key when reading the
# inline-format variant emitted by the legacy placement_engine.
_LEGACY_INLINE_KEY_RE = re.compile(
    r"({}):\s*".format("|".join(re.escape(k) for k in _LEGACY_FIELD_MAP))
)


class ElementLinkerError(Exception):
    pass


def _coerce_legacy_value(field_name, raw):
    if raw is None:
        return None
    raw = raw.strip()
    if raw == "" or raw == "Not found":
        return None
    if field_name in _INT_FIELDS:
        try:
            return int(raw)
        except (ValueError, TypeError):
            return None
    if field_name in _FLOAT_FIELDS:
        try:
            return float(raw)
        except (ValueError, TypeError):
            return None
    if field_name in _TUPLE_FIELDS:
        parts = [p.strip() for p in raw.split(",")]
        if len(parts) != 3:
            return None
        try:
            return [float(p) for p in parts]
        except (ValueError, TypeError):
            return None
    return raw


def _parse_legacy_kv(text):
    """Parse legacy text (multiline or inline) into ``{legacy_key: raw_string}``."""
    text = (text or "").strip()
    if not text:
        return {}
    if "\n" in text:
        out = {}
        for line in text.splitlines():
            if ":" not in line:
                continue
            key, _, value = line.partition(":")
            out[key.strip()] = value.strip()
        return out
    # Inline: anchor on known keys, slice between matches.
    matches = list(_LEGACY_INLINE_KEY_RE.finditer(text))
    if not matches:
        return {}
    out = {}
    for i, m in enumerate(matches):
        key = m.group(1)
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        value = text[start:end].rstrip().rstrip(",").strip(" ,")
        out[key] = value
    return out


class ElementLinker(object):
    """Typed wrapper around the JSON payload."""

    __slots__ = _FIELDS

    def __init__(self, **kwargs):
        for name in _FIELDS:
            setattr(self, name, kwargs.get(name))

    # --- factories ---------------------------------------------------

    @classmethod
    def from_dict(cls, d):
        if not isinstance(d, dict):
            raise ElementLinkerError(
                "Element_Linker JSON must be an object, got {}".format(type(d).__name__)
            )
        v = d.get("v")
        if v is not None and v != CODEC_VERSION:
            raise ElementLinkerError(
                "Unsupported Element_Linker codec version {} (expected {})".format(
                    v, CODEC_VERSION
                )
            )
        return cls(**{name: d.get(name) for name in _FIELDS})

    @classmethod
    def from_json(cls, text):
        if text is None:
            return None
        if not text.strip():
            return None
        try:
            d = json.loads(text)
        except (ValueError, TypeError) as exc:
            raise ElementLinkerError("Failed to parse Element_Linker JSON: {}".format(exc))
        return cls.from_dict(d)

    @classmethod
    def from_legacy_text(cls, text):
        """Read a legacy bespoke-text payload (multiline or inline)."""
        raw = _parse_legacy_kv(text)
        if not raw:
            return None
        fields = {}
        for legacy_key, value in raw.items():
            new_key = _LEGACY_FIELD_MAP.get(legacy_key)
            if new_key is None:
                continue
            if fields.get(new_key) is not None:
                # Earlier alt-spelling wins, mirroring legacy precedence.
                continue
            fields[new_key] = _coerce_legacy_value(new_key, value)
        return cls(**fields)

    # --- serialisation ----------------------------------------------

    def to_dict(self):
        out = {"v": CODEC_VERSION}
        for name in _FIELDS:
            out[name] = getattr(self, name)
        return out

    def to_json(self):
        return json.dumps(self.to_dict(), separators=(",", ":"), sort_keys=False)

    # --- conveniences -----------------------------------------------

    def __eq__(self, other):
        if not isinstance(other, ElementLinker):
            return NotImplemented
        return self.to_dict() == other.to_dict()

    def __ne__(self, other):
        result = self.__eq__(other)
        return result if result is NotImplemented else not result

    def __repr__(self):
        return "ElementLinker(led_id={!r}, set_id={!r}, element_id={!r})".format(
            self.led_id, self.set_id, self.element_id
        )
