# -*- coding: utf-8 -*-
"""
Reusable PipeSegment model and aggregation helpers.

All lengths/diameters are kept in Revit internal units (feet) unless a caller
converts them before assignment.
"""

import math
import re

try:
    from pyrevit import DB  # type: ignore
except Exception:
    DB = None


DEFAULT_PARAMETER_MAP = {
    "horizontal_length": ("CED_Horizontal Length", "Horizontal Length"),
    "vertical_length": ("CED_Vertical Length", "Vertical Length"),
    "trunk_parent": ("TrunkParent", "Trunk Parent"),
    "branch_child_1": ("BranchChild1", "Branch Child 1"),
    "branch_child_2": ("BranchChild2", "Branch Child 2"),
    "branch_children_csv": ("BranchChildren", "Branch Children"),
    "evaporation_capacity": ("Evaporation Capacity", "CED_Evaporation Capacity"),
    "identity_mark": ("Identity Mark", "Mark", "System ID", "SystemID", "System Id"),
    "is_leaf": ("is_leaf", "Is Leaf"),
    "is_root": ("is_root", "Is Root"),
    "radius": ("Radius", "Pipe Radius"),
    "diameter": ("Diameter", "Pipe Diameter"),
}


def _coerce_param_names(param_map, key):
    value = None
    if isinstance(param_map, dict):
        value = param_map.get(key)
    if value is None:
        value = DEFAULT_PARAMETER_MAP.get(key, ())
    if isinstance(value, (list, tuple)):
        return [v for v in value if v]
    if value:
        return [value]
    return []


def _get_builtin_parameter():
    if DB is None:
        return None
    return getattr(DB, "BuiltInParameter", None)


def _get_storage_type():
    if DB is None:
        return None
    return getattr(DB, "StorageType", None)


def _get_param_by_builtin(element, bip_name):
    if element is None or not bip_name:
        return None
    bip_type = _get_builtin_parameter()
    if bip_type is None:
        return None
    bip = getattr(bip_type, bip_name, None)
    if bip is None:
        return None
    try:
        return element.get_Parameter(bip)
    except Exception:
        return None


def _get_param_by_names(element, names):
    if element is None:
        return None
    for name in names or []:
        try:
            p = element.LookupParameter(name)
            if p is not None:
                return p
        except Exception:
            continue
    return None


def _as_string(param):
    if param is None:
        return None
    try:
        value = param.AsString()
        if value is not None:
            return value
    except Exception:
        pass
    try:
        value = param.AsValueString()
        if value is not None:
            return value
    except Exception:
        pass
    return None


def _as_text(param):
    text = _as_string(param)
    if text is not None and str(text).strip():
        return str(text).strip()
    if param is None:
        return None
    try:
        if not param.HasValue:
            return None
    except Exception:
        pass
    try:
        value = param.AsInteger()
        if value is not None:
            return str(value).strip()
    except Exception:
        pass
    try:
        value = param.AsDouble()
        if value is not None:
            return str(value).strip()
    except Exception:
        pass
    return None


def _as_double(param, default=0.0):
    if param is None:
        return float(default)
    try:
        if not param.HasValue:
            return float(default)
    except Exception:
        pass
    try:
        value = param.AsDouble()
        if value is not None:
            return float(value)
    except Exception:
        pass
    try:
        text = _as_string(param)
        if text is not None and str(text).strip():
            return float(str(text).strip())
    except Exception:
        pass
    return float(default)



def _safe_float(value, default=0.0):
    try:
        return float(value)
    except Exception:
        return float(default)


def _as_bool(param, default=None):
    if param is None:
        return default
    try:
        if not param.HasValue:
            return default
    except Exception:
        pass
    try:
        value = param.AsInteger()
        return bool(int(value))
    except Exception:
        pass
    text = _as_string(param)
    if text is None:
        return default
    lowered = str(text).strip().lower()
    if lowered in ("1", "true", "yes", "y"):
        return True
    if lowered in ("0", "false", "no", "n"):
        return False
    return default


def _set_param_value(param, value):
    if param is None:
        return False
    try:
        if param.IsReadOnly:
            return False
    except Exception:
        pass

    storage_type = _get_storage_type()
    try:
        if storage_type is not None and param.StorageType == storage_type.String:
            param.Set("" if value is None else str(value))
            return True
        if storage_type is not None and param.StorageType == storage_type.Integer:
            if isinstance(value, bool):
                param.Set(1 if value else 0)
            elif value is None:
                param.Set(0)
            else:
                param.Set(int(round(float(value))))
            return True
        if storage_type is not None and param.StorageType == storage_type.Double:
            param.Set(0.0 if value is None else float(value))
            return True
    except Exception:
        pass

    try:
        if isinstance(value, bool):
            param.Set(1 if value else 0)
            return True
    except Exception:
        pass
    try:
        if isinstance(value, (int, float)):
            param.Set(float(value))
            return True
    except Exception:
        pass
    try:
        param.Set("" if value is None else str(value))
        return True
    except Exception:
        return False


def _curve_component_lengths(pipe):
    try:
        location = getattr(pipe, "Location", None)
        curve = getattr(location, "Curve", None)
        if curve is None:
            return 0.0, 0.0
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        dx = float(p1.X - p0.X)
        dy = float(p1.Y - p0.Y)
        dz = float(p1.Z - p0.Z)
        horizontal = math.sqrt((dx * dx) + (dy * dy))
        vertical = abs(dz)
        # Classify by dominant direction (3x factor threshold).
        if vertical >= 3.0 * horizontal:
            return 0.0, vertical
        if horizontal >= 3.0 * vertical:
            return horizontal, 0.0
        return horizontal, vertical
    except Exception:
        return 0.0, 0.0


def _normalize_children(children):
    out = []
    for child in children or []:
        if child is None:
            continue
        text = str(child).strip()
        if not text:
            continue
        if text not in out:
            out.append(text)
    return out[:2]


def _split_children_csv(text):
    if not text:
        return []
    raw = str(text)
    for token in (";", "|"):
        raw = raw.replace(token, ",")
    return _normalize_children([part.strip() for part in raw.split(",")])


def _get_type_element(element):
    if element is None:
        return None
    try:
        doc = element.Document
        type_id = element.GetTypeId()
        if doc is None or type_id is None:
            return None
        return doc.GetElement(type_id)
    except Exception:
        return None


def _find_named_param_case_insensitive(element, names):
    if element is None:
        return None
    normalized = set()
    for n in names or []:
        if n is None:
            continue
        normalized.add(str(n).strip().lower())
    if not normalized:
        return None

    # Exact lookup first.
    direct = _get_param_by_names(element, names)
    if direct is not None:
        return direct

    # Then robust scan for case/spacing variants.
    try:
        for param in element.GetOrderedParameters():
            if param is None:
                continue
            definition = getattr(param, "Definition", None)
            name = getattr(definition, "Name", None) if definition is not None else None
            if name is None:
                continue
            if str(name).strip().lower() in normalized:
                return param
    except Exception:
        pass
    return None


def _get_first_nonempty_text_from_names(element, names, include_type=True):
    if element is None:
        return None
    candidates = [n for n in (names or []) if n]
    if not candidates:
        return None

    param = _find_named_param_case_insensitive(element, candidates)
    value = _as_text(param)
    if value:
        return value

    if include_type:
        type_element = _get_type_element(element)
        if type_element is not None:
            param = _find_named_param_case_insensitive(type_element, candidates)
            value = _as_text(param)
            if value:
                return value
    return None


def _get_identity_mark_text(pipe, parameter_map):
    custom_names = _coerce_param_names(parameter_map, "identity_mark")
    # Include common variants even if custom map overrides identity names.
    for fallback_name in ("Identity Mark", "Mark", "System ID", "SystemID", "System Id"):
        if fallback_name not in custom_names:
            custom_names.append(fallback_name)

    value = _get_first_nonempty_text_from_names(pipe, custom_names, include_type=True)
    if value:
        return value

    # Final fallback to built-in Mark.
    bip_param = _get_param_by_builtin(pipe, "ALL_MODEL_MARK")
    value = _as_text(bip_param)
    if value:
        return value
    return None


def _safe_element_id_int(element):
    try:
        return element.Id.IntegerValue
    except Exception:
        return None


def _iter_connectors(element):
    if element is None:
        return []
    connector_sets = []
    try:
        cm = getattr(element, "ConnectorManager", None)
        if cm is not None:
            connector_sets.append(getattr(cm, "Connectors", None))
    except Exception:
        pass
    try:
        mep_model = getattr(element, "MEPModel", None)
        if mep_model is not None:
            cm = getattr(mep_model, "ConnectorManager", None)
            if cm is not None:
                connector_sets.append(getattr(cm, "Connectors", None))
    except Exception:
        pass

    connectors = []
    for conn_set in connector_sets:
        if conn_set is None:
            continue
        try:
            for conn in conn_set:
                connectors.append(conn)
        except Exception:
            continue
    return connectors


def _flow_dir_name(connector):
    try:
        direction = connector.Direction
        return str(direction).upper()
    except Exception:
        return ""


def _is_flow_in(connector):
    name = _flow_dir_name(connector)
    return name.endswith("IN")


def _is_flow_out(connector):
    name = _flow_dir_name(connector)
    return name.endswith("OUT")


def _connected_owner_ids(connector):
    ids = []
    if connector is None:
        return ids
    try:
        for ref in connector.AllRefs:
            owner = getattr(ref, "Owner", None)
            owner_id = _safe_element_id_int(owner)
            if owner_id is not None:
                ids.append(owner_id)
    except Exception:
        pass
    return ids


class PipeSegment(object):
    def __init__(
        self,
        horizontal_length=0.0,
        vertical_length=0.0,
        trunk_parent=None,
        branch_children=None,
        evaporation_capacity=0.0,
        identity_mark=None,
        is_leaf=None,
        is_root=None,
        radius=0.0,
        diameter=0.0,
        source_element_id=None,
    ):
        self.horizontal_length = float(horizontal_length or 0.0)
        self.vertical_length = float(vertical_length or 0.0)
        self.trunk_parent = str(trunk_parent).strip() if trunk_parent is not None and str(trunk_parent).strip() else None
        self.branch_children = _normalize_children(branch_children or [])
        self.evaporation_capacity = float(evaporation_capacity or 0.0)
        self.identity_mark = str(identity_mark).strip() if identity_mark is not None and str(identity_mark).strip() else None
        self.radius = float(radius or 0.0)
        self.diameter = float(diameter or 0.0)
        self.source_element_id = source_element_id

        if self.radius <= 0.0 and self.diameter > 0.0:
            self.radius = self.diameter / 2.0
        if self.diameter <= 0.0 and self.radius > 0.0:
            self.diameter = self.radius * 2.0

        self.is_leaf = bool(is_leaf) if is_leaf is not None else (len(self.branch_children) == 0)
        self.is_root = bool(is_root) if is_root is not None else (self.trunk_parent is None)

    def update_topology_flags(self):
        self.is_leaf = len(self.branch_children) == 0
        self.is_root = self.trunk_parent is None

    def add_branch_child(self, child_identity_mark):
        child = str(child_identity_mark).strip() if child_identity_mark is not None else ""
        if not child:
            return False
        if child in self.branch_children:
            return False
        if len(self.branch_children) >= 2:
            raise ValueError("PipeSegment supports a maximum of 2 BranchChildren.")
        self.branch_children.append(child)
        self.update_topology_flags()
        return True

    def set_trunk_parent(self, parent_identity_mark):
        parent = str(parent_identity_mark).strip() if parent_identity_mark is not None else ""
        self.trunk_parent = parent if parent else None
        self.update_topology_flags()

    def set_trunk_parent_from_upstream_candidates(self, upstream_segments):
        candidates = [seg for seg in (upstream_segments or []) if isinstance(seg, PipeSegment)]
        if not candidates:
            self.trunk_parent = None
            self.update_topology_flags()
            return None
        chosen = sorted(candidates, key=lambda seg: float(seg.radius or 0.0), reverse=True)[0]
        self.trunk_parent = chosen.identity_mark
        self.update_topology_flags()
        return self.trunk_parent

    @classmethod
    def from_revit_pipe(cls, pipe, parameter_map=None):
        parameter_map = parameter_map or {}
        horizontal_length, vertical_length = _curve_component_lengths(pipe)

        identity_mark = _get_identity_mark_text(pipe, parameter_map)

        diameter_param = _get_param_by_builtin(pipe, "RBS_PIPE_DIAMETER_PARAM")
        if diameter_param is None:
            diameter_param = _get_param_by_names(pipe, _coerce_param_names(parameter_map, "diameter"))
        diameter = _as_double(diameter_param, default=0.0)

        radius_param = _get_param_by_names(pipe, _coerce_param_names(parameter_map, "radius"))
        radius = _as_double(radius_param, default=(diameter / 2.0 if diameter else 0.0))

        evap_param = _get_param_by_names(pipe, _coerce_param_names(parameter_map, "evaporation_capacity"))
        evaporation_capacity = _as_double(evap_param, default=0.0)

        trunk_param = _get_param_by_names(pipe, _coerce_param_names(parameter_map, "trunk_parent"))
        trunk_parent = _as_string(trunk_param)

        children = []
        child1_param = _get_param_by_names(pipe, _coerce_param_names(parameter_map, "branch_child_1"))
        child2_param = _get_param_by_names(pipe, _coerce_param_names(parameter_map, "branch_child_2"))
        csv_param = _get_param_by_names(pipe, _coerce_param_names(parameter_map, "branch_children_csv"))
        children.extend(_split_children_csv(_as_string(csv_param)))
        c1 = _as_string(child1_param)
        c2 = _as_string(child2_param)
        if c1:
            children.append(c1)
        if c2:
            children.append(c2)
        branch_children = _normalize_children(children)

        is_leaf_param = _get_param_by_names(pipe, _coerce_param_names(parameter_map, "is_leaf"))
        is_root_param = _get_param_by_names(pipe, _coerce_param_names(parameter_map, "is_root"))
        is_leaf = _as_bool(is_leaf_param, default=None)
        is_root = _as_bool(is_root_param, default=None)

        element_id = None
        try:
            element_id = pipe.Id.IntegerValue
        except Exception:
            element_id = None

        return cls(
            horizontal_length=horizontal_length,
            vertical_length=vertical_length,
            trunk_parent=trunk_parent,
            branch_children=branch_children,
            evaporation_capacity=evaporation_capacity,
            identity_mark=identity_mark,
            is_leaf=is_leaf,
            is_root=is_root,
            radius=radius,
            diameter=diameter,
            source_element_id=element_id,
        )

    def apply_to_revit_pipe(self, pipe, parameter_map=None):
        parameter_map = parameter_map or {}
        writes = {}

        # Write to explicit "Identity Mark" first if present, then built-in Mark.
        mark_param = _find_named_param_case_insensitive(pipe, _coerce_param_names(parameter_map, "identity_mark"))
        if mark_param is None:
            type_elem = _get_type_element(pipe)
            if type_elem is not None:
                mark_param = _find_named_param_case_insensitive(type_elem, _coerce_param_names(parameter_map, "identity_mark"))
        if mark_param is None:
            mark_param = _get_param_by_builtin(pipe, "ALL_MODEL_MARK")
        writes["identity_mark"] = _set_param_value(mark_param, self.identity_mark)

        field_specs = (
            ("horizontal_length", self.horizontal_length),
            ("vertical_length", self.vertical_length),
            ("trunk_parent", self.trunk_parent),
            ("evaporation_capacity", self.evaporation_capacity),
            ("is_leaf", self.is_leaf),
            ("is_root", self.is_root),
            ("radius", self.radius),
            ("diameter", self.diameter),
        )
        for field_name, value in field_specs:
            param = _get_param_by_names(pipe, _coerce_param_names(parameter_map, field_name))
            writes[field_name] = _set_param_value(param, value)

        child_1 = self.branch_children[0] if len(self.branch_children) > 0 else ""
        child_2 = self.branch_children[1] if len(self.branch_children) > 1 else ""
        child_csv = ",".join(self.branch_children)

        writes["branch_child_1"] = _set_param_value(
            _get_param_by_names(pipe, _coerce_param_names(parameter_map, "branch_child_1")),
            child_1,
        )
        writes["branch_child_2"] = _set_param_value(
            _get_param_by_names(pipe, _coerce_param_names(parameter_map, "branch_child_2")),
            child_2,
        )
        writes["branch_children_csv"] = _set_param_value(
            _get_param_by_names(pipe, _coerce_param_names(parameter_map, "branch_children_csv")),
            child_csv,
        )
        return writes

    def to_dict(self):
        return {
            "horizontal_length": self.horizontal_length,
            "vertical_length": self.vertical_length,
            "trunk_parent": self.trunk_parent,
            "branch_children": list(self.branch_children),
            "evaporation_capacity": self.evaporation_capacity,
            "identity_mark": self.identity_mark,
            "is_leaf": self.is_leaf,
            "is_root": self.is_root,
            "radius": self.radius,
            "diameter": self.diameter,
            "source_element_id": self.source_element_id,
        }

    @staticmethod
    def SumHorizontalPipeLengthPerID(pipe_segments):
        return SumHorizontalPipeLengthPerID(pipe_segments)

    @staticmethod
    def SumVerticalPipeLengthPerID(pipe_segments):
        return SumVerticalPipeLengthPerID(pipe_segments)

    @staticmethod
    def SumEvaporationCapacityPerID(pipe_segments):
        return SumEvaporationCapacityPerID(pipe_segments)

    @staticmethod
    def CheckEvaporationCapacitySums(pipe_segments, tolerance=1e-06):
        return CheckEvaporationCapacitySums(pipe_segments, tolerance=tolerance)

    @staticmethod
    def PrintSystemTotals(pipe_segments):
        return PrintPipeSegmentTotalsPerID(pipe_segments)

    @staticmethod
    def BuildFromRevitPipes(pipes, parameter_map=None, infer_relationships=True):
        return BuildPipeSegmentsFromRevitPipes(
            pipes,
            parameter_map=parameter_map,
            infer_relationships=infer_relationships,
        )


def _segment_id(segment):
    seg_id = getattr(segment, "identity_mark", None)
    if seg_id is None:
        return "<Unassigned ID>"
    text = str(seg_id).strip()
    return text if text else "<Unassigned ID>"


def _normalize_id_text(value):
    if value is None:
        return None
    text = str(value).strip()
    return text if text else None


_NATURAL_TOKEN_RE = re.compile(r"(\d+)")


def _natural_sort_key(value):
    text = str(value or "")
    parts = _NATURAL_TOKEN_RE.split(text)
    key = []
    for part in parts:
        if not part:
            continue
        if part.isdigit():
            key.append((0, int(part)))
        else:
            key.append((1, part.lower()))
    return key


def _append_unique(target_list, value):
    if value in target_list:
        return False
    target_list.append(value)
    return True


def _sum_per_id(pipe_segments, attr_name):
    totals = {}
    for segment in pipe_segments or []:
        if not isinstance(segment, PipeSegment):
            continue
        seg_id = _segment_id(segment)
        try:
            value = float(getattr(segment, attr_name, 0.0) or 0.0)
        except Exception:
            value = 0.0
        totals[seg_id] = totals.get(seg_id, 0.0) + value
    return totals


def OrderSystemIDsRootToLeaf(pipe_segments, allowed_ids=None, unassigned_label="<Unassigned ID>"):
    segments = [seg for seg in (pipe_segments or []) if isinstance(seg, PipeSegment)]
    nodes = set()
    if allowed_ids is not None:
        nodes = set([str(x) for x in allowed_ids if x is not None and str(x).strip()])
    else:
        for seg in segments:
            nodes.add(_segment_id(seg))

    if not nodes:
        return []

    adjacency = {}
    incoming = {}

    def _ensure_node(node_id):
        if node_id not in adjacency:
            adjacency[node_id] = []
        incoming.setdefault(node_id, 0)

    def _add_edge(parent_id, child_id):
        if not parent_id or not child_id:
            return
        if parent_id == child_id:
            return
        if parent_id not in nodes or child_id not in nodes:
            return
        _ensure_node(parent_id)
        _ensure_node(child_id)
        if _append_unique(adjacency[parent_id], child_id):
            incoming[child_id] = incoming.get(child_id, 0) + 1

    sorted_segments = sorted(
        segments,
        key=lambda seg: (
            _safe_float(getattr(seg, "source_element_id", None), 0.0),
            _segment_id(seg),
        ),
    )

    for seg in sorted_segments:
        seg_id = _segment_id(seg)
        if seg_id not in nodes:
            continue
        _ensure_node(seg_id)

        for child in getattr(seg, "branch_children", []) or []:
            child_id = _normalize_id_text(child)
            if child_id:
                _add_edge(seg_id, child_id)

    for seg in sorted_segments:
        child_id = _segment_id(seg)
        parent_id = _normalize_id_text(getattr(seg, "trunk_parent", None))
        if parent_id:
            _add_edge(parent_id, child_id)

    for node in list(nodes):
        _ensure_node(node)

    roots = [node for node in nodes if incoming.get(node, 0) == 0]
    roots = sorted(roots, key=_natural_sort_key)

    ordered = []
    visited = set()

    def _dfs(node_id, path):
        if node_id in path:
            return
        if node_id in visited:
            return
        visited.add(node_id)
        ordered.append(node_id)

        children = list(adjacency.get(node_id, []) or [])
        for child_id in children:
            _dfs(child_id, path | set([node_id]))

    for root_id in roots:
        _dfs(root_id, set())

    for node_id in sorted(nodes, key=_natural_sort_key):
        if node_id not in visited:
            _dfs(node_id, set())

    if unassigned_label in ordered:
        ordered = [x for x in ordered if x != unassigned_label] + [unassigned_label]
    return ordered


def SumHorizontalPipeLengthPerID(pipe_segments):
    return _sum_per_id(pipe_segments, "horizontal_length")


def SumVerticalPipeLengthPerID(pipe_segments):
    return _sum_per_id(pipe_segments, "vertical_length")


def SumEvaporationCapacityPerID(pipe_segments):
    return _sum_per_id(pipe_segments, "evaporation_capacity")


def CheckEvaporationCapacitySums(pipe_segments, tolerance=1e-06):
    segments = [seg for seg in (pipe_segments or []) if isinstance(seg, PipeSegment)]
    by_identity = {}
    for seg in segments:
        seg_id = _segment_id(seg)
        by_identity[seg_id] = seg

    mismatches = []
    for parent in segments:
        if not parent.branch_children:
            continue
        expected = float(parent.evaporation_capacity or 0.0)
        actual = 0.0
        missing = []
        for child_id in parent.branch_children:
            child_key = str(child_id).strip()
            child = by_identity.get(child_key)
            if child is None:
                missing.append(child_key)
                continue
            actual += float(child.evaporation_capacity or 0.0)
        if missing or abs(expected - actual) > float(tolerance):
            mismatches.append(
                {
                    "parent_id": _segment_id(parent),
                    "parent_evaporation_capacity": expected,
                    "children_evaporation_capacity_sum": actual,
                    "difference": expected - actual,
                    "missing_children": missing,
                }
            )

    return {"is_valid": len(mismatches) == 0, "mismatches": mismatches}


def BuildPipeSegmentsFromRevitPipes(pipes, parameter_map=None, infer_relationships=True):
    segments = []
    by_element_id = {}
    by_identity = {}
    source_pipes_by_id = {}

    for pipe in pipes or []:
        seg = PipeSegment.from_revit_pipe(pipe, parameter_map=parameter_map)
        segments.append(seg)
        element_id = seg.source_element_id
        if element_id is not None:
            by_element_id[element_id] = seg
            source_pipes_by_id[element_id] = pipe
        seg_id = _segment_id(seg)
        by_identity[seg_id] = seg

    if not infer_relationships:
        for seg in segments:
            seg.update_topology_flags()
        return segments

    parent_candidates = {}
    for element_id, pipe in source_pipes_by_id.items():
        this_segment = by_element_id.get(element_id)
        if this_segment is None:
            continue
        candidates = parent_candidates.setdefault(element_id, set())

        for conn in _iter_connectors(pipe):
            is_in = _is_flow_in(conn)
            is_out = _is_flow_out(conn)
            for owner_id in _connected_owner_ids(conn):
                if owner_id == element_id or owner_id not in by_element_id:
                    continue
                neighbor = by_element_id.get(owner_id)
                if neighbor is None:
                    continue
                if is_in:
                    candidates.add(owner_id)
                    continue
                if is_out:
                    continue
                if float(neighbor.radius or 0.0) >= float(this_segment.radius or 0.0):
                    candidates.add(owner_id)

    # Assign trunk parent based on largest radius upstream candidate.
    for element_id, candidate_ids in parent_candidates.items():
        segment = by_element_id.get(element_id)
        if segment is None:
            continue
        candidate_segments = [by_element_id.get(cid) for cid in candidate_ids if by_element_id.get(cid) is not None]
        if not candidate_segments:
            segment.trunk_parent = None
            segment.update_topology_flags()
            continue
        parent = sorted(candidate_segments, key=lambda seg: float(seg.radius or 0.0), reverse=True)[0]
        segment.trunk_parent = parent.identity_mark
        segment.update_topology_flags()

    # Build branch children lists from parent references.
    for seg in segments:
        seg.branch_children = []
    children_by_parent = {}
    for child in segments:
        parent_id = child.trunk_parent
        if not parent_id:
            continue
        children_by_parent.setdefault(str(parent_id).strip(), []).append(child)

    for parent_identity, children in children_by_parent.items():
        parent = by_identity.get(parent_identity)
        if parent is None:
            continue
        sorted_children = sorted(children, key=lambda seg: float(seg.radius or 0.0), reverse=True)
        parent.branch_children = _normalize_children([seg.identity_mark for seg in sorted_children if seg.identity_mark])
        parent.update_topology_flags()

    for seg in segments:
        seg.update_topology_flags()
    return segments


def PrintPipeSegmentTotalsPerID(pipe_segments, ordered_ids=None):
    vertical_totals = SumVerticalPipeLengthPerID(pipe_segments)
    horizontal_totals = SumHorizontalPipeLengthPerID(pipe_segments)
    evap_totals = SumEvaporationCapacityPerID(pipe_segments)

    all_ids = set(vertical_totals.keys()) | set(horizontal_totals.keys()) | set(evap_totals.keys())
    if ordered_ids is not None:
        ordered_ids = [sid for sid in ordered_ids if sid in all_ids]
        for sid in sorted(all_ids):
            if sid not in ordered_ids:
                ordered_ids.append(sid)
    else:
        ordered_ids = sorted(all_ids)
    lines = []
    for system_id in ordered_ids:
        line = (
            "System ID: {0} | Vertical Length Total: {1:.6f} | "
            "Horizontal Length Total: {2:.6f} | Evaporation Capacity Total: {3:.6f}"
        ).format(
            system_id,
            vertical_totals.get(system_id, 0.0),
            horizontal_totals.get(system_id, 0.0),
            evap_totals.get(system_id, 0.0),
        )
        print(line)
        lines.append(line)
    return lines


__all__ = [
    "PipeSegment",
    "DEFAULT_PARAMETER_MAP",
    "SumHorizontalPipeLengthPerID",
    "SumVerticalPipeLengthPerID",
    "SumEvaporationCapacityPerID",
    "CheckEvaporationCapacitySums",
    "OrderSystemIDsRootToLeaf",
    "BuildPipeSegmentsFromRevitPipes",
    "PrintPipeSegmentTotalsPerID",
]

