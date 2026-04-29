# -*- coding: utf-8 -*-
"""
Placement engine.

Replaces the six legacy placement tools (PlaceLinkedElements + 2 filtered
variants, PlaceCADElements + 2 filtered variants). Single entry point:

    targets = collect_targets(...)
    matches = match_targets(targets, profiles, mode)
    result = execute_placement(doc, matches, options)

Three target sources:
    * Linked Revit model — every FamilyInstance / Group in a chosen
      RevitLinkInstance, with parent_filter-style matching.
    * CSV file — rebased decimal coordinates produced upstream.
    * DWG link — best-effort block-instance walk (block names may be
      partially or fully unavailable depending on the DWG; the
      collector returns whatever it could extract).

Matching modes:
    * ``family_name_strip_suffix`` — for linked Revit. Strips trailing
      ``_NNN`` from both sides and compares family names case-insensitively.
    * ``cad_aliases`` — for CSV / DWG. Each profile may declare a
      comma-separated alias list under ``equipment_properties.cad_aliases``;
      a target whose name matches any alias is placed against that profile.
      Same trailing-``_NNN`` strip applies.
"""

import io
import math
import os
import re

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInParameter,
    ElementId,
    FamilyInstance,
    FamilySymbol,
    FilteredElementCollector,
    Group,
    GroupType,
    ImportInstance,
    LocationPoint,
    Options,
    RevitLinkInstance,
    Transform,
    XYZ,
)
from Autodesk.Revit.DB.Structure import StructuralType  # noqa: E402

import element_linker as _el
import element_linker_io as _el_io
import geometry
import hosted_annotations
import links


# ---------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------

CAD_ALIASES_KEY = "cad_aliases"

MATCH_FAMILY_NAME_STRIP_SUFFIX = "family_name_strip_suffix"
MATCH_CAD_ALIASES = "cad_aliases"

SOURCE_LINKED_REVIT = "linked_revit"
SOURCE_HOST_MODEL = "host_model"
SOURCE_CSV = "csv"
SOURCE_DWG_LINK = "dwg_link"
SOURCE_PICKED_POINT = "picked_point"


# ---------------------------------------------------------------------
# Pure-logic helpers (testable offline)
# ---------------------------------------------------------------------

_TRAILING_SUFFIX_RE = re.compile(r"_\d+$")


def strip_trailing_suffix(name):
    """Drop a trailing ``_NNN`` (one or more digits) from ``name``.

    ``"AC_BLOCK"`` -> ``"AC_BLOCK"``
    ``"AC_BLOCK_1"`` -> ``"AC_BLOCK"``
    ``"AC_BLOCK_42"`` -> ``"AC_BLOCK"``
    ``"AC_BLOCK_2A"`` -> ``"AC_BLOCK_2A"`` (only stripped if all digits)
    """
    if not name:
        return ""
    return _TRAILING_SUFFIX_RE.sub("", str(name))


def normalize_name(name):
    """Lowercase + strip trailing ``_NNN`` for case-insensitive matching.

    Whitespace is trimmed *before* the suffix strip so that a trailing
    space doesn't hide the suffix from the regex.
    """
    return strip_trailing_suffix((name or "").strip()).lower()


def collect_profile_aliases(profile):
    """Return the set of CAD aliases declared on ``profile``.

    Reads ``equipment_properties[cad_aliases]``. Value can be a list of
    strings or a single comma-separated string. Empty / missing -> set().
    """
    if not isinstance(profile, dict):
        return set()
    props = profile.get("equipment_properties") or {}
    if not isinstance(props, dict):
        return set()
    raw = props.get(CAD_ALIASES_KEY)
    if raw is None:
        return set()
    items = []
    if isinstance(raw, list):
        items = [str(x) for x in raw if x is not None]
    elif isinstance(raw, str):
        items = [s for s in raw.split(",")]
    else:
        items = [str(raw)]
    out = set()
    for item in items:
        norm = normalize_name(item)
        if norm:
            out.add(norm)
    return out


def profile_family_names(profile):
    """Return the set of family-name keys we'll match against for linked
    Revit. Pulls from ``parent_filter.family_name_pattern`` and from the
    profile's own name ("Family : Type") split."""
    if not isinstance(profile, dict):
        return set()
    out = set()
    pf = profile.get("parent_filter") or {}
    if isinstance(pf, dict):
        fam = pf.get("family_name_pattern")
        if fam:
            out.add(normalize_name(fam))
    name = profile.get("name") or ""
    if " : " in name:
        fam, _ = name.split(" : ", 1)
        if fam:
            out.add(normalize_name(fam))
    return {n for n in out if n}


def _profile_id_label(profile):
    return "{}  ({})".format(
        profile.get("name") or "(unnamed)",
        profile.get("id") or "?",
    )


# ---------------------------------------------------------------------
# Target dataclass-equivalent (kept Py2/3 compatible)
# ---------------------------------------------------------------------

class Target(object):
    """One anchor location for placement.

    ``source``      one of SOURCE_*
    ``name``        block name / family name / csv name
    ``world_pt``    (x, y, z) in the host doc's coordinate frame, feet
    ``rotation_deg`` rotation around Z, host frame, degrees
    ``level_id``    ElementId.Value (int) for placement level, or None
    ``link_inst``   RevitLinkInstance if linked-Revit, else None
    ``link_elem_id`` int ElementId of the linked element, else None
    """

    __slots__ = (
        "source", "name", "world_pt", "rotation_deg",
        "level_id", "link_inst", "link_elem_id",
    )

    def __init__(self, source, name, world_pt, rotation_deg=0.0,
                 level_id=None, link_inst=None, link_elem_id=None):
        self.source = source
        self.name = name or ""
        self.world_pt = tuple(world_pt) if world_pt else (0.0, 0.0, 0.0)
        self.rotation_deg = float(rotation_deg or 0.0)
        self.level_id = level_id
        self.link_inst = link_inst
        self.link_elem_id = link_elem_id

    def __repr__(self):
        return "Target(source={}, name={!r}, world_pt={})".format(
            self.source, self.name, self.world_pt
        )


class Match(object):
    """One (target, profile) pair the user can choose to place or skip."""

    __slots__ = ("target", "profile", "skip")

    def __init__(self, target, profile, skip=False):
        self.target = target
        self.profile = profile
        self.skip = skip


# ---------------------------------------------------------------------
# Matching
# ---------------------------------------------------------------------

def _match_one_linked_revit(target, profiles):
    """A linked-Revit target's ``name`` is the element's family name.
    Match against each profile's family_name set, both stripped."""
    target_key = normalize_name(target.name)
    if not target_key:
        return []
    out = []
    for p in profiles:
        if target_key in profile_family_names(p):
            out.append(p)
    return out


def _match_one_cad(target, profiles):
    """A CAD target's ``name`` is the block name. Match against each
    profile's ``cad_aliases`` (also stripped)."""
    target_key = normalize_name(target.name)
    if not target_key:
        return []
    out = []
    for p in profiles:
        if target_key in collect_profile_aliases(p):
            out.append(p)
    return out


def match_targets(targets, profiles, mode):
    """Cross-product match: every (target, profile) pair where the
    matcher accepts becomes a ``Match``. One target may match multiple
    profiles (via overlapping aliases / multiple family-name keys)."""
    matcher = (
        _match_one_linked_revit if mode == MATCH_FAMILY_NAME_STRIP_SUFFIX
        else _match_one_cad
    )
    out = []
    for t in targets:
        for p in matcher(t, profiles):
            out.append(Match(t, p))
    return out


# ---------------------------------------------------------------------
# Skip-already-placed
# ---------------------------------------------------------------------

def _placed_set_anchor_signatures(doc):
    """Return a set of ``(set_id, anchor_signature)`` for every placed
    fixture in the active document. Anchor signature is a coarse
    rounded-XY tuple — used to skip re-placing onto an anchor that
    already has the profile present."""
    signatures = set()
    for klass in (FamilyInstance, Group):
        collector = FilteredElementCollector(doc).OfClass(klass).WhereElementIsNotElementType()
        for elem in collector:
            linker = _el_io.read_from_element(elem)
            if linker is None or not linker.set_id:
                continue
            anchor_pt = linker.parent_location_ft
            if anchor_pt and len(anchor_pt) >= 2:
                key = (
                    linker.set_id,
                    round(float(anchor_pt[0]), 1),
                    round(float(anchor_pt[1]), 1),
                )
                signatures.add(key)
    return signatures


def filter_already_placed(doc, matches):
    """Drop matches whose target+profile pair already has a placement.

    Returns ``(kept, skipped_count)``.
    """
    sigs = _placed_set_anchor_signatures(doc)
    kept = []
    skipped = 0
    for m in matches:
        target_set_ids = [
            s.get("id") for s in (m.profile.get("linked_sets") or [])
            if isinstance(s, dict) and s.get("id")
        ]
        target_xy = (round(m.target.world_pt[0], 1), round(m.target.world_pt[1], 1))
        already = any(
            (sid, target_xy[0], target_xy[1]) in sigs
            for sid in target_set_ids
        )
        if already:
            skipped += 1
            continue
        kept.append(m)
    return kept, skipped


# ---------------------------------------------------------------------
# Target collection: linked Revit
# ---------------------------------------------------------------------

def collect_linked_revit_link_instances(doc):
    """Return ``[RevitLinkInstance, ...]`` for every loaded Revit link."""
    out = []
    for link_inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        if link_inst is None:
            continue
        try:
            if link_inst.GetLinkDocument() is None:
                continue
        except Exception:
            continue
        out.append(link_inst)
    return out


def find_targets_in_host_model(doc):
    """Walk the host doc's FamilyInstances + Groups and produce
    one Target per element. Anchors are matched by family name
    (same suffix-strip rules as linked Revit)."""
    if doc is None:
        return []
    out = []
    identity = Transform.Identity
    for klass in (FamilyInstance, Group):
        collector = (
            FilteredElementCollector(doc)
            .OfClass(klass)
            .WhereElementIsNotElementType()
        )
        for elem in collector:
            name = _element_family_name(elem)
            if not name:
                continue
            pt = _element_location_point(elem)
            if pt is None:
                continue
            rot_deg = _element_rotation_deg(elem, identity)
            eid = elem.Id
            eid_val = getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)
            out.append(Target(
                source=SOURCE_HOST_MODEL,
                name=name,
                world_pt=(pt.X, pt.Y, pt.Z),
                rotation_deg=rot_deg,
                level_id=None,
                link_inst=None,
                link_elem_id=eid_val,
            ))
    return out


def find_targets_in_linked_revit(link_inst):
    """Walk the linked doc's FamilyInstances + Groups and produce
    one Target per element."""
    if link_inst is None:
        return []
    link_doc = link_inst.GetLinkDocument()
    if link_doc is None:
        return []
    transform = links.get_link_transform(link_inst) or Transform.Identity
    out = []
    for klass in (FamilyInstance, Group):
        collector = (
            FilteredElementCollector(link_doc)
            .OfClass(klass)
            .WhereElementIsNotElementType()
        )
        for elem in collector:
            name = _element_family_name(elem)
            if not name:
                continue
            pt = _element_location_point(elem)
            if pt is None:
                continue
            world_pt = transform.OfPoint(pt)
            rot_deg = _element_rotation_deg(elem, transform)
            eid = elem.Id
            eid_val = getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)
            out.append(Target(
                source=SOURCE_LINKED_REVIT,
                name=name,
                world_pt=(world_pt.X, world_pt.Y, world_pt.Z),
                rotation_deg=rot_deg,
                level_id=None,
                link_inst=link_inst,
                link_elem_id=eid_val,
            ))
    return out


def _element_family_name(elem):
    if isinstance(elem, FamilyInstance):
        sym = getattr(elem, "Symbol", None)
        if sym is not None:
            family = getattr(sym, "Family", None)
            return family.Name if family else ""
    if isinstance(elem, Group):
        gtype = getattr(elem, "GroupType", None)
        return gtype.Name if gtype else ""
    return ""


def _element_location_point(elem):
    loc = getattr(elem, "Location", None)
    if loc is None:
        return None
    pt = getattr(loc, "Point", None)
    if pt is not None:
        return pt
    bbox = elem.get_BoundingBox(None)
    if bbox is not None:
        return XYZ(
            (bbox.Min.X + bbox.Max.X) / 2.0,
            (bbox.Min.Y + bbox.Max.Y) / 2.0,
            (bbox.Min.Z + bbox.Max.Z) / 2.0,
        )
    return None


def _element_rotation_deg(elem, transform):
    """Note: ``hasattr(loc, 'Rotation')`` would invoke the property
    getter in pythonnet, which throws InvalidOperationException for
    line-based families etc. Use ``isinstance(loc, LocationPoint)``
    instead — that's a safe type check that doesn't touch the property.
    """
    rad = 0.0
    loc = getattr(elem, "Location", None)
    if isinstance(loc, LocationPoint):
        try:
            rad = loc.Rotation
        except Exception:
            rad = 0.0
    if transform is not None:
        try:
            local = XYZ(math.cos(rad), math.sin(rad), 0.0)
            v = transform.OfVector(local)
            return geometry.normalize_angle(math.degrees(math.atan2(v.Y, v.X)))
        except Exception:
            pass
    return geometry.normalize_angle(math.degrees(rad))


# ---------------------------------------------------------------------
# Target collection: CSV
# ---------------------------------------------------------------------

def collect_csv_target_files():
    """No-op stub: caller picks the file via a UI dialog."""
    return []


_FT_IN_RE = re.compile(
    r"""^\s*
        (?P<sign>-)?\s*
        (?:(?P<feet>\d+)\s*'\s*)?
        (?:-?\s*(?P<whole>\d+))?
        (?:\s*(?P<num>\d+)\s*/\s*(?P<den>\d+))?
        \s*\"?\s*$
    """,
    re.VERBOSE,
)


def parse_feet_value(value):
    """Parse a length value to decimal feet.

    Accepts:
        ``"12.5"``                  -> 12.5
        ``"12'-6\""``               -> 12.5
        ``"12'-6 1/2\""``           -> 12.5417
        ``"-3'-2\""``               -> -3.1667
        ``"-1.25"``                 -> -1.25

    Returns ``None`` on parse failure.
    """
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    # Try plain decimal first.
    try:
        return float(text)
    except (TypeError, ValueError):
        pass
    m = _FT_IN_RE.match(text)
    if not m:
        return None
    sign = -1.0 if m.group("sign") else 1.0
    feet = float(m.group("feet") or 0)
    whole_in = float(m.group("whole") or 0)
    num = m.group("num")
    den = m.group("den")
    frac_in = (float(num) / float(den)) if (num and den and float(den) != 0) else 0.0
    inches = whole_in + frac_in
    if feet == 0 and whole_in == 0 and frac_in == 0:
        return None
    return sign * (feet + inches / 12.0)


_HEADER_NAME_KEYS = ("name", "block_name", "block", "blockname", "family")
_HEADER_X_KEYS = ("x", "position x", "position_x", "pos_x", "px")
_HEADER_Y_KEYS = ("y", "position y", "position_y", "pos_y", "py")
_HEADER_Z_KEYS = ("z", "position z", "position_z", "pos_z", "pz", "elevation")
_HEADER_ROT_KEYS = ("rotation", "rot", "rotation_deg", "angle", "rotation (deg)")


def _resolve_header(headers, keys):
    lower = [(h or "").strip().lower() for h in headers]
    for key in keys:
        if key in lower:
            return lower.index(key)
    return None


def find_targets_in_csv(path):
    """Read ``path`` as CSV. Headers are auto-detected; the file must
    have at least a name column and X / Y columns. Z and rotation are
    optional (default 0)."""
    if not path or not os.path.isfile(path):
        return []
    import csv
    rows = []
    with io.open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        for raw in reader:
            rows.append(raw)
    if not rows:
        return []
    headers = rows[0]
    name_idx = _resolve_header(headers, _HEADER_NAME_KEYS)
    x_idx = _resolve_header(headers, _HEADER_X_KEYS)
    y_idx = _resolve_header(headers, _HEADER_Y_KEYS)
    z_idx = _resolve_header(headers, _HEADER_Z_KEYS)
    rot_idx = _resolve_header(headers, _HEADER_ROT_KEYS)
    if name_idx is None or x_idx is None or y_idx is None:
        raise CsvParseError(
            "CSV missing required columns. Need at least a name column "
            "(name/block_name) and X/Y columns. Headers seen: {}".format(headers)
        )
    targets = []
    for ridx, raw in enumerate(rows[1:], start=2):
        if not raw:
            continue
        try:
            name = (raw[name_idx] or "").strip()
        except IndexError:
            continue
        if not name:
            continue
        x = parse_feet_value(raw[x_idx] if x_idx < len(raw) else None)
        y = parse_feet_value(raw[y_idx] if y_idx < len(raw) else None)
        z = parse_feet_value(raw[z_idx] if z_idx is not None and z_idx < len(raw) else None) or 0.0
        rot = parse_feet_value(raw[rot_idx] if rot_idx is not None and rot_idx < len(raw) else None) or 0.0
        if x is None or y is None:
            continue
        targets.append(Target(
            source=SOURCE_CSV,
            name=name,
            world_pt=(x, y, z),
            rotation_deg=rot,
            level_id=None,
        ))
    return targets


class CsvParseError(Exception):
    pass


# ---------------------------------------------------------------------
# Target collection: DWG link (best-effort)
# ---------------------------------------------------------------------

def collect_dwg_link_instances(doc):
    """Return loaded DWG-link ImportInstances. Imports (not links) are
    excluded — they can't be reliably differentiated from blocks."""
    out = []
    for inst in FilteredElementCollector(doc).OfClass(ImportInstance):
        if inst is None:
            continue
        try:
            if not inst.IsLinked:
                continue
        except Exception:
            pass
        out.append(inst)
    return out


def find_targets_in_dwg_link(import_inst):
    """Best-effort walk of a DWG link's geometry. Revit's API does not
    reliably expose individual block names, so this may return zero
    usable targets — when it does, names are derived from the geometry's
    own ``Symbol.Name`` if available.
    """
    if import_inst is None:
        return []
    out = []
    try:
        opts = Options()
        opts.IncludeNonVisibleObjects = False
        opts.ComputeReferences = False
        geom = import_inst.get_Geometry(opts)
    except Exception:
        return []
    base_transform = import_inst.GetTotalTransform()
    for obj in geom or []:
        if obj is None:
            continue
        # Each insert / nested xref shows up as a GeometryInstance.
        sym = getattr(obj, "Symbol", None)
        try:
            tx = obj.Transform
        except Exception:
            tx = None
        name = ""
        if sym is not None:
            name = getattr(sym, "Name", "") or ""
        if not name:
            # Some DWGs surface block names via a category mapping; try
            # the GraphicsStyle (a per-layer record). This is unreliable
            # but better than nothing.
            try:
                gs = obj.GraphicsStyleId
                gs_elem = import_inst.Document.GetElement(gs) if gs else None
                if gs_elem is not None:
                    name = getattr(gs_elem, "Name", "") or ""
            except Exception:
                pass
        if not name:
            continue
        if tx is None:
            world_pt = (base_transform.Origin.X, base_transform.Origin.Y, base_transform.Origin.Z)
            rot_deg = 0.0
        else:
            origin = base_transform.OfPoint(tx.Origin)
            world_pt = (origin.X, origin.Y, origin.Z)
            try:
                bx = base_transform.OfVector(tx.BasisX)
                rot_deg = geometry.normalize_angle(math.degrees(math.atan2(bx.Y, bx.X)))
            except Exception:
                rot_deg = 0.0
        out.append(Target(
            source=SOURCE_DWG_LINK,
            name=name,
            world_pt=world_pt,
            rotation_deg=rot_deg,
            level_id=None,
        ))
    return out


# ---------------------------------------------------------------------
# Placement execution
# ---------------------------------------------------------------------

class PlacementResult(object):
    def __init__(self):
        self.placed_fixture_count = 0
        self.placed_annotation_count = 0
        self.element_linker_writes = 0
        self.warnings = []
        self.errors = []
        self.skipped_already_placed = 0


class PlacementOptions(object):
    """Knobs the dialog feeds into ``execute_placement``."""

    def __init__(self,
                 skip_already_placed=True,
                 default_level_id=None,
                 transaction_action="Place from CAD or Linked Model"):
        self.skip_already_placed = skip_already_placed
        self.default_level_id = default_level_id
        self.transaction_action = transaction_action


def _activate_symbol(symbol):
    if symbol is None:
        return
    try:
        if not symbol.IsActive:
            symbol.Activate()
    except Exception:
        pass


def _find_family_symbol(doc, family_name, type_name):
    """Find a FamilySymbol matching ``family_name : type_name``."""
    for sym in FilteredElementCollector(doc).OfClass(FamilySymbol):
        family = getattr(sym, "Family", None)
        if family is None:
            continue
        if family.Name != family_name:
            continue
        if sym.Name == type_name:
            return sym
    return None


def _find_group_type(doc, group_name):
    for gt in FilteredElementCollector(doc).OfClass(GroupType):
        if gt.Name == group_name:
            return gt
    return None


def _split_label(label):
    """``"Family : Type"`` -> ``("Family", "Type")``. Single-word labels
    fall back to ``(label, "")``."""
    if not label:
        return "", ""
    if " : " in label:
        family, type_name = label.split(" : ", 1)
        return family.strip(), type_name.strip()
    return label.strip(), ""


def _place_fixture(doc, led, anchor_world_pt, anchor_rotation_deg, level_id):
    """Place one LED at the resolved world point. Returns the placed
    Element or None if placement failed."""
    label = led.get("label") or ""
    is_group = bool(led.get("is_group"))
    offsets_list = led.get("offsets") or []
    offset = offsets_list[0] if offsets_list else {
        "x_inches": 0.0, "y_inches": 0.0, "z_inches": 0.0, "rotation_deg": 0.0,
    }
    target_pt_t = geometry.target_point_from_offsets(
        anchor_world_pt, anchor_rotation_deg, offset
    )
    target_rot_deg = geometry.child_rotation_from_offsets(
        anchor_rotation_deg, offset
    )
    target_pt = XYZ(target_pt_t[0], target_pt_t[1], target_pt_t[2])

    if is_group:
        group_type = _find_group_type(doc, label)
        if group_type is None:
            return None
        group = doc.Create.PlaceGroup(target_pt, group_type)
        if abs(target_rot_deg) > geometry.Tolerances.ROTATION_DEG:
            try:
                from Autodesk.Revit.DB import ElementTransformUtils, Line
                axis = Line.CreateBound(target_pt, XYZ(target_pt.X, target_pt.Y, target_pt.Z + 1.0))
                ElementTransformUtils.RotateElement(
                    doc, group.Id, axis, math.radians(target_rot_deg)
                )
            except Exception:
                pass
        return group

    family_name, type_name = _split_label(label)
    if not family_name:
        return None
    symbol = _find_family_symbol(doc, family_name, type_name)
    if symbol is None:
        return None
    _activate_symbol(symbol)
    level = doc.GetElement(ElementId(int(level_id))) if level_id else None
    try:
        if level is not None:
            inst = doc.Create.NewFamilyInstance(
                target_pt, symbol, level, StructuralType.NonStructural
            )
        else:
            inst = doc.Create.NewFamilyInstance(
                target_pt, symbol, StructuralType.NonStructural
            )
    except Exception:
        try:
            inst = doc.Create.NewFamilyInstance(
                target_pt, symbol, StructuralType.NonStructural
            )
        except Exception:
            return None
    if inst is not None and abs(target_rot_deg) > geometry.Tolerances.ROTATION_DEG:
        try:
            from Autodesk.Revit.DB import ElementTransformUtils, Line
            axis = Line.CreateBound(target_pt, XYZ(target_pt.X, target_pt.Y, target_pt.Z + 1.0))
            ElementTransformUtils.RotateElement(
                doc, inst.Id, axis, math.radians(target_rot_deg)
            )
        except Exception:
            pass
    return inst


def _write_linker(elem, led, profile, target):
    """Stamp the placed element with an Element_Linker payload so audit
    and re-placement tools can find it."""
    if elem is None:
        return False
    set_id = None
    for s in profile.get("linked_sets") or []:
        if isinstance(s, dict):
            set_id = s.get("id")
            break
    parent_elem_id = (
        target.link_elem_id
        if target.source in (SOURCE_LINKED_REVIT, SOURCE_HOST_MODEL)
        else None
    )
    payload = _el.ElementLinker(
        led_id=led.get("id"),
        set_id=set_id,
        location_ft=list(_element_world_point(elem)) if elem is not None else None,
        rotation_deg=_element_world_rotation_deg(elem),
        parent_rotation_deg=target.rotation_deg,
        parent_element_id=parent_elem_id,
        level_id=_element_level_id_value(elem),
        element_id=_element_id_value(elem),
        facing=_element_facing(elem),
        host_name=target.name,
        parent_location_ft=list(target.world_pt),
    )
    try:
        _el_io.write_to_element(elem, payload)
        return True
    except _el_io.ElementLinkerIOError:
        return False


def _element_id_value(elem):
    if elem is None:
        return None
    eid = elem.Id
    return getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)


def _element_level_id_value(elem):
    lid = getattr(elem, "LevelId", None)
    if lid is None:
        try:
            param = elem.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM)
            if param is not None:
                lid = param.AsElementId()
        except Exception:
            return None
    if lid is None:
        return None
    return getattr(lid, "Value", None) or getattr(lid, "IntegerValue", None)


def _element_world_point(elem):
    pt = _element_location_point(elem)
    if pt is None:
        return (0.0, 0.0, 0.0)
    return (pt.X, pt.Y, pt.Z)


def _element_world_rotation_deg(elem):
    rad = 0.0
    loc = getattr(elem, "Location", None)
    if isinstance(loc, LocationPoint):
        try:
            rad = loc.Rotation
        except Exception:
            rad = 0.0
    return geometry.normalize_angle(math.degrees(rad))


def _element_facing(elem):
    facing = getattr(elem, "FacingOrientation", None)
    if facing is None:
        return None
    return [facing.X, facing.Y, facing.Z]


def execute_placement(doc, matches, options=None):
    """Place every non-skipped match. Caller manages the transaction."""
    if options is None:
        options = PlacementOptions()
    result = PlacementResult()

    if options.skip_already_placed:
        kept, skipped = filter_already_placed(doc, [m for m in matches if not m.skip])
        result.skipped_already_placed = skipped
    else:
        kept = [m for m in matches if not m.skip]

    for m in kept:
        anchor = m.target.world_pt
        anchor_rot = m.target.rotation_deg
        for set_dict in m.profile.get("linked_sets") or []:
            if not isinstance(set_dict, dict):
                continue
            for led in set_dict.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                placed = _place_fixture(
                    doc, led, anchor, anchor_rot, options.default_level_id
                )
                if placed is None:
                    result.warnings.append(
                        "Could not place LED {} ({}) — family/group not loaded.".format(
                            led.get("label") or "?", led.get("id") or "?"
                        )
                    )
                    continue
                result.placed_fixture_count += 1
                if _write_linker(placed, led, m.profile, m.target):
                    result.element_linker_writes += 1

    return result
