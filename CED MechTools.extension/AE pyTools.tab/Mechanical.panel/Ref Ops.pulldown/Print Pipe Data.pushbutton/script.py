# -*- coding: utf-8 -*-
__title__ = "Print Pipe Data"
__doc__ = "Collect pipes by selected worksets, summarize PipeSegment totals, and export to Excel."

import csv
import io
import math
import os
import re
import sys
from datetime import datetime

import System
from System import Activator, Type
from System.Reflection import BindingFlags
from System.Runtime.InteropServices import Marshal
from pyrevit import DB, forms, revit, script

from LogicClasses.PipeSegment import (
    BuildPipeSegmentsFromRevitPipes,
    PrintPipeSegmentTotalsPerID,
    SumEvaporationCapacityPerID,
    SumHorizontalPipeLengthPerID,
    SumVerticalPipeLengthPerID,
)

logger = script.get_logger()
output = script.get_output()
doc = revit.doc

try:
    if sys.getrecursionlimit() < 20000:
        sys.setrecursionlimit(20000)
except Exception:
    pass

_PIPE_CATEGORY_NAMES = (
    "OST_PipeCurves",
    "OST_FlexPipeCurves",
    "OST_PipePlaceholder",
    "OST_FabricationPipework",
)

_IDENTITY_PARAM_NAMES = (
    "Identity Mark",
    "Mark",
    "System ID",
    "SystemID",
    "System Id",
)

_TRUNK_ID_RE = re.compile(r"^\s*(\d+)([A-Za-z])\s*$")
_DECIMAL_ID_RE = re.compile(r"^\s*(\d+)\.(\d+)([A-Za-z])\s*$")
_R_BRANCH_ID_RE = re.compile(r"^\s*R([A-Za-z])(\d+)\s*$", re.IGNORECASE)


def _safe_float(value, default=0.0):
    try:
        return float(value)
    except Exception:
        return float(default)


def _safe_text(value):
    if value is None:
        return None
    text = str(value).strip()
    return text if text else None


def _natural_sort_key(value):
    text = str(value or "")
    key = []
    token = ""
    token_is_digit = None

    for ch in text:
        is_digit = ch.isdigit()
        if token_is_digit is None:
            token = ch
            token_is_digit = is_digit
            continue
        if is_digit == token_is_digit:
            token += ch
            continue
        if token_is_digit:
            key.append((0, int(token)))
        else:
            key.append((1, token.lower()))
        token = ch
        token_is_digit = is_digit

    if token:
        if token_is_digit:
            key.append((0, int(token)))
        else:
            key.append((1, token.lower()))
    return key


def _letter_rank(letter):
    if not letter:
        return 999
    try:
        return ord(str(letter).strip().lower()[0]) - ord("a")
    except Exception:
        return 999


def _parse_id_shape(system_id):
    text = _safe_text(system_id) or ""
    if text == "<Unassigned ID>":
        return ("unassigned",)

    m = _TRUNK_ID_RE.match(text)
    if m:
        return ("trunk", int(m.group(1)), _letter_rank(m.group(2)))

    m = _R_BRANCH_ID_RE.match(text)
    if m:
        return ("rbranch", _letter_rank(m.group(1)), int(m.group(2)))

    m = _DECIMAL_ID_RE.match(text)
    if m:
        return ("decimal", int(m.group(1)), int(m.group(2)), _letter_rank(m.group(3)))

    return ("other", _natural_sort_key(text))


def _root_rank(system_id):
    shape = _parse_id_shape(system_id)
    kind = shape[0]
    if kind == "trunk":
        return (0, shape[2], 0, shape[1], _natural_sort_key(system_id))
    if kind == "decimal":
        return (0, shape[3], 1, shape[1], shape[2], _natural_sort_key(system_id))
    if kind == "rbranch":
        return (0, shape[1], 2, shape[2], _natural_sort_key(system_id))
    if kind == "other":
        return (1, 999, 0, _natural_sort_key(system_id))
    if kind == "unassigned":
        return (9, 999, 0, _natural_sort_key(system_id))
    return (2, 999, 0, _natural_sort_key(system_id))


def _branch_sort_key_for_trunk(node_id, trunk_id, id_first_idx):
    node_shape = _parse_id_shape(node_id)
    trunk_shape = _parse_id_shape(trunk_id)
    trunk_letter = trunk_shape[2] if trunk_shape and trunk_shape[0] == "trunk" else 999
    trunk_number = trunk_shape[1] if trunk_shape and trunk_shape[0] == "trunk" else 999999
    kind = node_shape[0]

    if kind == "rbranch":
        same_letter = 0 if node_shape[1] == trunk_letter else 1
        return (0, same_letter, node_shape[2], id_first_idx.get(node_id, 10 ** 9), _natural_sort_key(node_id))

    if kind == "decimal":
        same_letter = 0 if node_shape[3] == trunk_letter else 1
        same_number = 0 if node_shape[1] == trunk_number else 1
        return (1, same_letter, same_number, node_shape[2], id_first_idx.get(node_id, 10 ** 9), _natural_sort_key(node_id))

    if kind == "other":
        return (2, id_first_idx.get(node_id, 10 ** 9), _natural_sort_key(node_id))

    if kind == "unassigned":
        return (9, id_first_idx.get(node_id, 10 ** 9), _natural_sort_key(node_id))

    return (3, id_first_idx.get(node_id, 10 ** 9), _natural_sort_key(node_id))



def _fast_order_system_ids(pipe_elements, identity_by_pipe_id, pipe_segments):
    keys = set()

    for pipe in (pipe_elements or []):
        try:
            pid = pipe.Id.IntegerValue
        except Exception:
            continue
        sid = _safe_text(identity_by_pipe_id.get(pid)) or "<Unassigned ID>"
        keys.add(sid)

    for seg in (pipe_segments or []):
        sid = _safe_text(getattr(seg, "identity_mark", None)) or "<Unassigned ID>"
        keys.add(sid)

    ordered = sorted(keys, key=lambda sid: (_root_rank(sid), _natural_sort_key(sid)))
    if "<Unassigned ID>" in ordered:
        ordered = [x for x in ordered if x != "<Unassigned ID>"] + ["<Unassigned ID>"]
    return ordered

def _args_array(*args):
    return System.Array[System.Object](list(args))


def _set(obj, prop, value):
    obj.GetType().InvokeMember(prop, BindingFlags.SetProperty, None, obj, _args_array(value))


def _get(obj, prop):
    try:
        return obj.GetType().InvokeMember(prop, BindingFlags.GetProperty, None, obj, None)
    except Exception:
        return None


def _call(obj, name, *args):
    t = obj.GetType()
    try:
        return t.InvokeMember(name, BindingFlags.InvokeMethod, None, obj, _args_array(*args))
    except Exception:
        try:
            return t.InvokeMember(name, BindingFlags.GetProperty, None, obj, _args_array(*args) if args else None)
        except Exception:
            return None


def _param_to_text(param):
    if param is None:
        return None
    try:
        if hasattr(param, "HasValue") and not param.HasValue:
            return None
    except Exception:
        pass

    for reader in ("AsString", "AsValueString"):
        try:
            value = getattr(param, reader)()
            text = _safe_text(value)
            if text:
                return text
        except Exception:
            continue

    for reader in ("AsInteger", "AsDouble"):
        try:
            value = getattr(param, reader)()
            text = _safe_text(value)
            if text:
                return text
        except Exception:
            continue

    return None


def _direct_identity_mark(pipe):
    if pipe is None:
        return None

    for name in _IDENTITY_PARAM_NAMES:
        try:
            param = pipe.LookupParameter(name)
        except Exception:
            param = None
        text = _param_to_text(param)
        if text:
            return text

    try:
        param = pipe.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
    except Exception:
        param = None
    text = _param_to_text(param)
    if text:
        return text

    target_names = set([str(n).strip().lower() for n in _IDENTITY_PARAM_NAMES])
    try:
        for param in pipe.GetOrderedParameters():
            if param is None:
                continue
            definition = getattr(param, "Definition", None)
            pname = getattr(definition, "Name", None) if definition is not None else None
            pname_key = _safe_text(pname)
            if not pname_key:
                continue
            if pname_key.lower() in target_names:
                text = _param_to_text(param)
                if text:
                    return text
    except Exception:
        pass

    return None


class _SystemTypeOption(forms.TemplateListItem):
    @property
    def name(self):
        return self.item or "<Unknown System Type>"


def _get_pipe_system_type_name(pipe):
    try:
        param = pipe.get_Parameter(DB.BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)
        if param is not None:
            type_id = param.AsElementId()
            if type_id is not None:
                sys_type_elem = doc.GetElement(type_id)
                if sys_type_elem is not None:
                    return sys_type_elem.Name
    except Exception:
        pass
    try:
        param = pipe.get_Parameter(DB.BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM)
        if param is not None:
            value = param.AsString() or param.AsValueString()
            if value:
                return value
    except Exception:
        pass
    return None


def _collect_all_pipes():
    categories = _collect_pipe_categories()
    pipe_elems = []
    seen = set()
    for bic in categories:
        collector = DB.FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
        for elem in collector:
            try:
                elem_id = elem.Id.IntegerValue
            except Exception:
                continue
            if elem_id in seen:
                continue
            seen.add(elem_id)
            pipe_elems.append(elem)
    return pipe_elems


def _collect_system_type_names(pipe_elements):
    names = set()
    for pipe in pipe_elements:
        name = _get_pipe_system_type_name(pipe)
        if name:
            names.add(name)
    return sorted(names, key=lambda n: (n or "").lower())


def _prompt_system_types(system_type_names):
    options = [_SystemTypeOption(name, checked=True) for name in system_type_names]
    selected = forms.SelectFromList.show(
        options,
        title="Select Pipe System Types",
        button_name="Collect Pipe Data",
        multiselect=True,
        return_all=True,
    )
    if selected is None:
        return None

    picked = []
    for option in selected:
        try:
            is_checked = bool(option)
        except Exception:
            is_checked = False
        if not is_checked:
            continue
        name = getattr(option, "item", None)
        if name is not None:
            picked.append(name)
    return picked


def _filter_pipes_by_system_types(pipe_elements, selected_type_names):
    name_set = set(selected_type_names)
    return [p for p in pipe_elements if _get_pipe_system_type_name(p) in name_set]


def _collect_pipe_categories():
    out = []
    for name in _PIPE_CATEGORY_NAMES:
        try:
            bic = getattr(DB.BuiltInCategory, name)
        except Exception:
            bic = None
        if bic is not None:
            out.append(bic)
    return out




def _iter_connectors(element):
    connectors = []
    if element is None:
        return connectors

    try:
        cm = getattr(element, "ConnectorManager", None)
        conn_set = getattr(cm, "Connectors", None) if cm is not None else None
        if conn_set is not None:
            for conn in conn_set:
                connectors.append(conn)
    except Exception:
        pass

    try:
        mep_model = getattr(element, "MEPModel", None)
        cm = getattr(mep_model, "ConnectorManager", None) if mep_model is not None else None
        conn_set = getattr(cm, "Connectors", None) if cm is not None else None
        if conn_set is not None:
            for conn in conn_set:
                connectors.append(conn)
    except Exception:
        pass

    return connectors


def _connected_pipe_ids(pipe, pipe_ids, fitting_pipe_cache=None, max_fitting_depth=4):
    out = set()
    if pipe is None:
        return out

    try:
        this_id = pipe.Id.IntegerValue
    except Exception:
        return out

    def _is_fitting_like(elem):
        if elem is None:
            return False
        try:
            eid = elem.Id.IntegerValue
        except Exception:
            return False
        if eid in pipe_ids:
            return False

        try:
            connectors = _iter_connectors(elem)
            ccount = len(connectors)
        except Exception:
            return False

        if ccount < 2 or ccount > 4:
            return False

        try:
            cat = getattr(elem, "Category", None)
            cat_name = (getattr(cat, "Name", "") or "").strip().lower()
            if "equipment" in cat_name or "fixture" in cat_name or "terminal" in cat_name:
                return False
            if "fitting" in cat_name:
                return True
        except Exception:
            pass

        # Fallback: small connector-count non-pipe elements are usually fitting-like.
        return True

    def _is_inline_fitting(elem):
        if not _is_fitting_like(elem):
            return False
        try:
            ccount = len(_iter_connectors(elem))
        except Exception:
            return False
        return ccount == 2
    def _pipes_through_fitting(start_fitting):
        if start_fitting is None:
            return set()

        try:
            start_id = start_fitting.Id.IntegerValue
        except Exception:
            return set()

        if fitting_pipe_cache is not None and start_id in fitting_pipe_cache:
            return set(fitting_pipe_cache[start_id])

        pipes = set()
        visited = set()
        stack = [(start_fitting, 0)]

        while stack:
            fitting, depth = stack.pop()
            if fitting is None:
                continue
            try:
                fid = fitting.Id.IntegerValue
            except Exception:
                continue
            if fid in visited:
                continue
            visited.add(fid)

            for conn in _iter_connectors(fitting):
                try:
                    refs = conn.AllRefs
                except Exception:
                    refs = []
                for ref in refs:
                    try:
                        owner = ref.Owner
                        owner_id = owner.Id.IntegerValue
                    except Exception:
                        continue

                    if owner_id in pipe_ids:
                        pipes.add(owner_id)
                        continue

                    if owner_id == fid:
                        continue

                    if depth + 1 > max_fitting_depth:
                        continue

                    if _is_inline_fitting(owner):
                        stack.append((owner, depth + 1))

        if fitting_pipe_cache is not None:
            fitting_pipe_cache[start_id] = set(pipes)

        return pipes

    for conn in _iter_connectors(pipe):
        try:
            refs = conn.AllRefs
        except Exception:
            refs = []

        for ref in refs:
            try:
                owner = ref.Owner
                owner_id = owner.Id.IntegerValue
            except Exception:
                continue

            if owner_id == this_id:
                continue

            if owner_id in pipe_ids:
                out.add(owner_id)
                continue

            if not _is_fitting_like(owner):
                continue

            for pid in _pipes_through_fitting(owner):
                if pid != this_id and pid in pipe_ids:
                    out.add(pid)

    return out


def _order_system_ids_by_connectivity(pipe_elements, identity_by_pipe_id, pipe_segments):
    pipe_by_id = {}
    sid_by_pipe = {}
    sid_to_pipes = {}
    sid_set = set()
    first_idx_by_sid = {}
    radius_by_pid = {}
    sid_max_radius = {}

    for seg in (pipe_segments or []):
        seg_sid = _safe_text(getattr(seg, "identity_mark", None)) or "<Unassigned ID>"
        seg_radius = _safe_float(getattr(seg, "radius", 0.0), 0.0)
        sid_max_radius[seg_sid] = max(_safe_float(sid_max_radius.get(seg_sid, 0.0), 0.0), seg_radius)
        try:
            seg_pid = int(getattr(seg, "source_element_id", None))
        except Exception:
            seg_pid = None
        if seg_pid is not None:
            radius_by_pid[seg_pid] = max(_safe_float(radius_by_pid.get(seg_pid, 0.0), 0.0), seg_radius)

    for idx, pipe in enumerate(pipe_elements or []):
        try:
            pid = pipe.Id.IntegerValue
        except Exception:
            continue
        sid = _safe_text(identity_by_pipe_id.get(pid)) or "<Unassigned ID>"
        pipe_by_id[pid] = pipe
        sid_by_pipe[pid] = sid
        sid_to_pipes.setdefault(sid, set()).add(pid)
        sid_set.add(sid)
        pipe_radius = _safe_float(radius_by_pid.get(pid, 0.0), 0.0)
        sid_max_radius[sid] = max(_safe_float(sid_max_radius.get(sid, 0.0), 0.0), pipe_radius)
        if sid not in first_idx_by_sid:
            first_idx_by_sid[sid] = idx

    for seg in (pipe_segments or []):
        sid = _safe_text(getattr(seg, "identity_mark", None)) or "<Unassigned ID>"
        sid_set.add(sid)
        if sid not in first_idx_by_sid:
            first_idx_by_sid[sid] = 10 ** 9

    if not sid_set:
        return []

    pipe_ids = set(pipe_by_id.keys())

    # Narrow pipe-level graph (local only) used for proximity fallback.
    graph = {}
    fitting_cache = {}
    for pid, pipe in pipe_by_id.items():
        neighbors = _connected_pipe_ids(
            pipe,
            pipe_ids,
            fitting_pipe_cache=fitting_cache,
            max_fitting_depth=1,
        )
        graph.setdefault(pid, set())
        for npid in neighbors:
            if npid == pid or npid not in pipe_ids:
                continue
            graph[pid].add(npid)
            graph.setdefault(npid, set()).add(pid)

    def _sid_letter_rank(sid):
        sh = _parse_id_shape(sid)
        if sh[0] == "trunk":
            return sh[2]
        if sh[0] == "decimal":
            return sh[3]
        if sh[0] == "rbranch":
            return sh[1]
        return 999

    def _decimal_suffix_text(sid):
        text = _safe_text(sid) or ""
        m = _DECIMAL_ID_RE.match(text)
        if m is None:
            return None
        return m.group(2)

    def _decimal_depth(sid):
        suffix = _decimal_suffix_text(sid)
        return len(suffix) if suffix else 0

    def _decimal_parent_from_name(sid):
        text = _safe_text(sid) or ""
        m = _DECIMAL_ID_RE.match(text)
        if m is None:
            return None
        root_num = int(m.group(1))
        suffix = m.group(2)
        letter = m.group(3).upper()

        probe = suffix[:-1]
        while probe:
            cand = "{}.{}{}".format(root_num, probe, letter)
            if cand in sid_set:
                return cand
            probe = probe[:-1]

        trunk_cand = "{}{}".format(root_num, letter)
        if trunk_cand in sid_set:
            return trunk_cand
        return None

    def _is_branch_kind(sid):
        kind = _parse_id_shape(sid)[0]
        return kind in ("trunk", "decimal", "other")

    def _branch_pref_key(sid):
        sh = _parse_id_shape(sid)
        kind = sh[0]
        sid_radius = _safe_float(sid_max_radius.get(sid, 0.0), 0.0)

        # For RA parent assignment, first prefer larger-radius upstream candidates
        # at a shared fitting/tee, then prefer shallower branches.
        if kind == "trunk":
            kind_rank = 0
            depth_rank = 0
        elif kind == "decimal":
            kind_rank = 1
            depth_rank = _decimal_depth(sid)
        else:
            kind_rank = 2
            depth_rank = 999

        return (
            -sid_radius,
            kind_rank,
            depth_rank,
            first_idx_by_sid.get(sid, 10 ** 9),
            _natural_sort_key(sid),
        )

    def _lowest_trunk_for_letter(letter_rank):
        trunks = [sid for sid in sid_set if _parse_id_shape(sid)[0] == "trunk" and _sid_letter_rank(sid) == letter_rank]
        if not trunks:
            return None
        trunks = sorted(
            trunks,
            key=lambda sid: (_parse_id_shape(sid)[1], first_idx_by_sid.get(sid, 10 ** 9), _natural_sort_key(sid)),
        )
        return trunks[0]

    parent_by_sid = {}

    # Build non-RA branch hierarchy from naming.
    decimal_ids = [sid for sid in sid_set if _parse_id_shape(sid)[0] == "decimal"]
    decimal_ids = sorted(
        decimal_ids,
        key=lambda sid: (_decimal_depth(sid), first_idx_by_sid.get(sid, 10 ** 9), _natural_sort_key(sid)),
    )
    for sid in decimal_ids:
        parent_sid = _decimal_parent_from_name(sid)
        if parent_sid and parent_sid != sid:
            parent_by_sid[sid] = parent_sid

    branch_ids_by_letter = {}
    for sid in sid_set:
        if not _is_branch_kind(sid):
            continue
        branch_ids_by_letter.setdefault(_sid_letter_rank(sid), set()).add(sid)

    def _ranked_parent_candidates_for_ra(ra_sid):
        starts = list(sid_to_pipes.get(ra_sid, set()) or [])
        if not starts:
            return []

        letter = _sid_letter_rank(ra_sid)
        candidates = set(branch_ids_by_letter.get(letter, set()) or set())
        if not candidates:
            return []

        ordered = []
        def _is_local_fitting_owner(owner):
            if owner is None:
                return False
            try:
                oid = owner.Id.IntegerValue
            except Exception:
                return False
            if oid in pipe_ids:
                return False
            try:
                ccount = len(_iter_connectors(owner))
            except Exception:
                return False
            if ccount < 2 or ccount > 4:
                return False
            try:
                cat = getattr(owner, "Category", None)
                cat_name = (getattr(cat, "Name", "") or "").strip().lower()
                if "equipment" in cat_name or "fixture" in cat_name or "terminal" in cat_name:
                    return False
            except Exception:
                pass
            return True

        # 1) Strong local signal: direct connector and same-fitting neighbors.
        local_scores = {}
        for spid in starts:
            pipe = pipe_by_id.get(spid)
            if pipe is None:
                continue
            for conn in _iter_connectors(pipe):
                try:
                    refs = conn.AllRefs
                except Exception:
                    refs = []

                for ref in refs:
                    try:
                        owner = ref.Owner
                        owner_id = owner.Id.IntegerValue
                    except Exception:
                        continue

                    if owner_id == spid:
                        continue

                    if owner_id in pipe_ids:
                        cand_sid = sid_by_pipe.get(owner_id)
                        if cand_sid in candidates:
                            local_scores[cand_sid] = local_scores.get(cand_sid, 0.0) + 5.0
                        continue

                    if not _is_local_fitting_owner(owner):
                        continue

                    # Owner is fitting/equipment: collect pipes on this same fitting only.
                    fitting_pipe_ids = set()
                    for fconn in _iter_connectors(owner):
                        try:
                            frefs = fconn.AllRefs
                        except Exception:
                            frefs = []
                        for fref in frefs:
                            try:
                                fowner = fref.Owner
                                fowner_id = fowner.Id.IntegerValue
                            except Exception:
                                continue
                            if fowner_id in pipe_ids and fowner_id != spid:
                                fitting_pipe_ids.add(fowner_id)

                    for fp_id in fitting_pipe_ids:
                        cand_sid = sid_by_pipe.get(fp_id)
                        if cand_sid in candidates:
                            local_scores[cand_sid] = local_scores.get(cand_sid, 0.0) + 3.0

        if local_scores:
            ranked_local = sorted(
                list(local_scores.keys()),
                key=lambda sid: (
                    -_safe_float(local_scores.get(sid, 0.0), 0.0),
                    _branch_pref_key(sid),
                ),
            )
            for sid in ranked_local:
                if sid not in ordered:
                    ordered.append(sid)

        # 2) Immediate graph neighbors.
        graph_scores = {}
        for spid in starts:
            for npid in (graph.get(spid, set()) or set()):
                nsid = sid_by_pipe.get(npid)
                if nsid in candidates:
                    graph_scores[nsid] = graph_scores.get(nsid, 0.0) + 1.0

        if graph_scores:
            ranked_graph = sorted(
                list(graph_scores.keys()),
                key=lambda sid: (
                    -_safe_float(graph_scores.get(sid, 0.0), 0.0),
                    _branch_pref_key(sid),
                ),
            )
            for sid in ranked_graph:
                if sid not in ordered:
                    ordered.append(sid)

        # 3) Short BFS fallback on pipe graph.
        dist = {}
        queue = []
        for spid in starts:
            dist[spid] = 0
            queue.append(spid)

        hit_dist = {}
        head = 0
        while head < len(queue):
            pid = queue[head]
            head += 1
            d = dist.get(pid, 0)
            if d > 10:
                continue

            sid = sid_by_pipe.get(pid)
            if sid in candidates and sid not in hit_dist:
                hit_dist[sid] = d

            for npid in (graph.get(pid, set()) or set()):
                if npid not in dist:
                    dist[npid] = d + 1
                    queue.append(npid)

        if hit_dist:
            ranked_bfs = sorted(
                list(hit_dist.keys()),
                key=lambda sid: (
                    _safe_float(hit_dist.get(sid, 10 ** 9), 10 ** 9),
                    _branch_pref_key(sid),
                ),
            )
            for sid in ranked_bfs:
                if sid not in ordered:
                    ordered.append(sid)

        return ordered

    # Assign each RA leaf to its best local parent directly.
    rbranch_ids = [sid for sid in sid_set if _parse_id_shape(sid)[0] == "rbranch"]
    rbranch_ids = sorted(rbranch_ids, key=lambda sid: (first_idx_by_sid.get(sid, 10 ** 9), _natural_sort_key(sid)))

    for ra_sid in rbranch_ids:
        ranked = _ranked_parent_candidates_for_ra(ra_sid)
        picked = ranked[0] if ranked else None

        if picked and picked != ra_sid:
            parent_by_sid[ra_sid] = picked
    # Build recursive tree output.
    children_by_sid = {}
    for sid in sid_set:
        children_by_sid[sid] = []
    for child_sid, parent_sid in parent_by_sid.items():
        children_by_sid.setdefault(parent_sid, []).append(child_sid)

    def _root_key(sid):
        sh = _parse_id_shape(sid)
        kind = sh[0]
        idx = first_idx_by_sid.get(sid, 10 ** 9)
        if kind == "trunk":
            return (0, sh[2], sh[1], idx, _natural_sort_key(sid))
        if kind == "decimal":
            return (1, sh[3], sh[1], _decimal_depth(sid), idx, _natural_sort_key(sid))
        if kind == "rbranch":
            return (2, sh[1], sh[2], idx, _natural_sort_key(sid))
        if kind == "other":
            return (3, idx, _natural_sort_key(sid))
        return (9, idx, _natural_sort_key(sid))

    def _child_sort_key(parent_sid, child_sid):
        sh = _parse_id_shape(child_sid)
        kind = sh[0]
        idx = first_idx_by_sid.get(child_sid, 10 ** 9)

        if kind == "rbranch":
            return (0, sh[2], idx, _natural_sort_key(child_sid))

        if kind == "decimal":
            parent_suffix = _decimal_suffix_text(parent_sid) or ""
            child_suffix = _decimal_suffix_text(child_sid) or ""
            tail = child_suffix[len(parent_suffix):] if child_suffix.startswith(parent_suffix) else child_suffix
            try:
                tail_num = int(tail) if tail else 0
            except Exception:
                tail_num = 10 ** 9
            return (1, tail_num, len(child_suffix), idx, _natural_sort_key(child_sid))

        if kind == "trunk":
            return (2, sh[1], idx, _natural_sort_key(child_sid))

        return (3, idx, _natural_sort_key(child_sid))

    roots = [sid for sid in sid_set if sid not in parent_by_sid]
    roots = sorted(roots, key=_root_key)

    ordered = []
    visited = set()

    def _walk(node_sid):
        if node_sid in visited:
            return
        visited.add(node_sid)
        ordered.append(node_sid)

        children = sorted(list(children_by_sid.get(node_sid, []) or []), key=lambda sid: _child_sort_key(node_sid, sid))

        # Hard rule: emit all leaf children (R* IDs) before any branch child.
        leaf_children = []
        branch_children = []
        for child_sid in children:
            if _parse_id_shape(child_sid)[0] == "rbranch":
                leaf_children.append(child_sid)
            else:
                branch_children.append(child_sid)

        for child_sid in leaf_children:
            _walk(child_sid)
        for child_sid in branch_children:
            _walk(child_sid)

    for root_sid in roots:
        _walk(root_sid)

    for sid in sorted(list(sid_set), key=_root_key):
        if sid not in visited:
            ordered.append(sid)

    if "<Unassigned ID>" in ordered:
        ordered = [x for x in ordered if x != "<Unassigned ID>"] + ["<Unassigned ID>"]

    return ordered

def _order_ids_from_topology_segments(topo_segments, allowed_ids=None):
    segments = list(topo_segments or [])

    nodes = set()
    if allowed_ids is not None:
        for raw in (allowed_ids or []):
            sid = _safe_text(raw) or ("<Unassigned ID>" if str(raw or "").strip() == "<Unassigned ID>" else None)
            if sid:
                nodes.add(sid)
    else:
        for seg in segments:
            sid = _safe_text(getattr(seg, "identity_mark", None)) or "<Unassigned ID>"
            nodes.add(sid)

    if not nodes:
        return []

    adjacency = {}
    incoming = {}

    def _ensure(node_id):
        if node_id not in adjacency:
            adjacency[node_id] = []
        incoming.setdefault(node_id, 0)

    def _add_edge(parent_id, child_id):
        parent_sid = _safe_text(parent_id)
        child_sid = _safe_text(child_id)
        if not parent_sid or not child_sid:
            return
        if parent_sid == child_sid:
            return
        if parent_sid not in nodes or child_sid not in nodes:
            return
        _ensure(parent_sid)
        _ensure(child_sid)
        if child_sid in adjacency[parent_sid]:
            return
        adjacency[parent_sid].append(child_sid)
        incoming[child_sid] = incoming.get(child_sid, 0) + 1

    sorted_segments = sorted(
        segments,
        key=lambda seg: (
            _safe_float(getattr(seg, "source_element_id", None), 0.0),
            _natural_sort_key(_safe_text(getattr(seg, "identity_mark", None)) or "<Unassigned ID>"),
        ),
    )

    for seg in sorted_segments:
        sid = _safe_text(getattr(seg, "identity_mark", None)) or "<Unassigned ID>"
        if sid not in nodes:
            continue
        _ensure(sid)
        for child in (getattr(seg, "branch_children", None) or []):
            _add_edge(sid, child)

    for seg in sorted_segments:
        sid = _safe_text(getattr(seg, "identity_mark", None)) or "<Unassigned ID>"
        parent_sid = _safe_text(getattr(seg, "trunk_parent", None))
        if parent_sid:
            _add_edge(parent_sid, sid)

    for node in list(nodes):
        _ensure(node)

    def _child_key(child_sid):
        sh = _parse_id_shape(child_sid)
        kind = sh[0]
        if kind == "rbranch":
            return (0, sh[2], _natural_sort_key(child_sid))
        if kind == "decimal":
            return (1, sh[1], sh[2], _natural_sort_key(child_sid))
        if kind == "trunk":
            return (2, sh[1], _natural_sort_key(child_sid))
        if kind == "other":
            return (3, _natural_sort_key(child_sid))
        return (9, _natural_sort_key(child_sid))

    roots = [node for node in nodes if incoming.get(node, 0) == 0]
    roots = sorted(roots, key=lambda sid: (_root_rank(sid), _natural_sort_key(sid)))

    ordered = []
    visited = set()

    for root_sid in roots:
        stack = [root_sid]
        while stack:
            sid = stack.pop()
            if sid in visited:
                continue
            visited.add(sid)
            ordered.append(sid)

            children = sorted(list(adjacency.get(sid, []) or []), key=_child_key)
            for child_sid in reversed(children):
                if child_sid not in visited:
                    stack.append(child_sid)

    for sid in sorted(list(nodes), key=lambda s: (_root_rank(s), _natural_sort_key(s))):
        if sid not in visited:
            ordered.append(sid)

    if "<Unassigned ID>" in ordered:
        ordered = [x for x in ordered if x != "<Unassigned ID>"] + ["<Unassigned ID>"]

    return ordered

def _rows_from_totals(pipe_segments, ordered_ids=None):
    vertical = SumVerticalPipeLengthPerID(pipe_segments)
    horizontal = SumHorizontalPipeLengthPerID(pipe_segments)
    evap = SumEvaporationCapacityPerID(pipe_segments)
    keys = set(vertical.keys()) | set(horizontal.keys()) | set(evap.keys())

    ordered_keys = []
    if ordered_ids:
        ordered_keys = [sid for sid in ordered_ids if sid in keys]
        for sid in [(_safe_text(getattr(seg, "identity_mark", None)) or "<Unassigned ID>") for seg in (pipe_segments or [])]:
            if sid in keys and sid not in ordered_keys:
                ordered_keys.append(sid)

    if not ordered_keys:
        # last fallback only
        ordered_keys = sorted(keys, key=_natural_sort_key)

    _MIN_LENGTH = 0.5 / 12.0  # 0.5 inches in feet (Revit internal units)

    def _threshold(value):
        v = _safe_float(value, 0.0)
        if v < _MIN_LENGTH:
            return 0.0
        return float(math.ceil(v))

    rows = []
    for sid in ordered_keys:
        rows.append({
            "System ID": sid,
            "Vertical Length Total": _threshold(vertical.get(sid, 0.0)),
            "Horizontal Length Total": _threshold(horizontal.get(sid, 0.0)),
            "Evaporation Capacity Total": _safe_float(evap.get(sid, 0.0), 0.0),
        })
    return rows


def _default_export_name():
    title = (doc.Title or "Model").replace(".rvt", "").strip()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return "Pipe_Data_{}_{}.xlsx".format(title, stamp)


def _ensure_xlsx_path(path):
    if not path:
        return path
    if path.lower().endswith(".xlsx"):
        return path
    return path + ".xlsx"


def _write_excel_xlsx(path, rows):
    excel = workbooks = workbook = worksheets = worksheet = cells = None
    try:
        excel_type = Type.GetTypeFromProgID("Excel.Application")
        if excel_type is None:
            raise RuntimeError("Excel is not available on this machine.")

        excel = Activator.CreateInstance(excel_type)
        _set(excel, "Visible", False)
        _set(excel, "DisplayAlerts", False)

        workbooks = _get(excel, "Workbooks")
        workbook = _call(workbooks, "Add")
        worksheets = _get(workbook, "Worksheets")
        worksheet = _call(worksheets, "Item", 1)
        _set(worksheet, "Name", "Pipe Totals")
        cells = _get(worksheet, "Cells")

        headers = [
            "System ID",
            "Vertical Length Total",
            "Horizontal Length Total",
            "Evaporation Capacity Total",
        ]
        for col_idx, header in enumerate(headers, 1):
            cell = _call(cells, "Item", 1, col_idx)
            _set(cell, "Value2", header)

        for row_idx, row in enumerate(rows, 2):
            cell = _call(cells, "Item", row_idx, 1)
            _set(cell, "Value2", row.get("System ID", ""))
            cell = _call(cells, "Item", row_idx, 2)
            _set(cell, "Value2", _safe_float(row.get("Vertical Length Total", 0.0), 0.0))
            cell = _call(cells, "Item", row_idx, 3)
            _set(cell, "Value2", _safe_float(row.get("Horizontal Length Total", 0.0), 0.0))
            cell = _call(cells, "Item", row_idx, 4)
            _set(cell, "Value2", _safe_float(row.get("Evaporation Capacity Total", 0.0), 0.0))

        _call(_get(worksheet, "Columns"), "AutoFit")
        _call(workbook, "SaveAs", path)
    finally:
        try:
            if workbook:
                _call(workbook, "Close", False)
        except Exception:
            pass
        try:
            if excel:
                _call(excel, "Quit")
        except Exception:
            pass

        for obj in (cells, worksheet, worksheets, workbook, workbooks, excel):
            try:
                if obj:
                    Marshal.ReleaseComObject(obj)
            except Exception:
                pass


def _write_csv(path, rows):
    headers = [
        "System ID",
        "Vertical Length Total",
        "Horizontal Length Total",
        "Evaporation Capacity Total",
    ]
    with io.open(path, "w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers)
        writer.writeheader()
        for row in rows:
            writer.writerow({
                "System ID": row.get("System ID", ""),
                "Vertical Length Total": "{:.6f}".format(_safe_float(row.get("Vertical Length Total", 0.0), 0.0)),
                "Horizontal Length Total": "{:.6f}".format(_safe_float(row.get("Horizontal Length Total", 0.0), 0.0)),
                "Evaporation Capacity Total": "{:.6f}".format(_safe_float(row.get("Evaporation Capacity Total", 0.0), 0.0)),
            })


def main():
    if doc is None:
        forms.alert("No active Revit document.", title=__title__, exitscript=True)
    if doc.IsFamilyDocument:
        forms.alert("This tool requires a project document.", title=__title__, exitscript=True)

    all_pipe_elements = _collect_all_pipes()
    if not all_pipe_elements:
        forms.alert("No pipes were found in this model.", title=__title__, exitscript=True)

    system_type_names = _collect_system_type_names(all_pipe_elements)
    if not system_type_names:
        forms.alert("No system types were found on the pipe elements.", title=__title__, exitscript=True)

    selected_type_names = _prompt_system_types(system_type_names)
    if selected_type_names is None:
        return
    if not selected_type_names:
        forms.alert("No system types were selected.", title=__title__, exitscript=True)

    pipe_elements = _filter_pipes_by_system_types(all_pipe_elements, selected_type_names)
    if not pipe_elements:
        forms.alert(
            "No pipes were found for the selected system types:\n{}".format(", ".join(sorted(selected_type_names))),
            title=__title__,
            exitscript=True,
        )

    identity_by_pipe_id = {}
    for pipe in pipe_elements:
        try:
            pipe_id = pipe.Id.IntegerValue
        except Exception:
            continue
        identity_by_pipe_id[pipe_id] = _direct_identity_mark(pipe)

    pipe_segments = BuildPipeSegmentsFromRevitPipes(pipe_elements, infer_relationships=False)

    # Force identity keys from direct Revit reads.
    for seg in pipe_segments:
        try:
            seg_pid = int(seg.source_element_id)
        except Exception:
            seg_pid = None
        if seg_pid is None:
            continue
        sid = _safe_text(identity_by_pipe_id.get(seg_pid))
        if sid:
            seg.identity_mark = sid

    ordered_ids = _order_system_ids_by_connectivity(pipe_elements, identity_by_pipe_id, pipe_segments)

    _ = PrintPipeSegmentTotalsPerID(pipe_segments, ordered_ids=ordered_ids)

    rows = _rows_from_totals(pipe_segments, ordered_ids=ordered_ids)
    if not rows:
        forms.alert("No pipe totals were generated.", title=__title__, exitscript=True)

    unassigned_rows = [r for r in rows if str(r.get("System ID", "")).strip() == "<Unassigned ID>"]
    assigned_rows = [r for r in rows if str(r.get("System ID", "")).strip() != "<Unassigned ID>"]

    direct_assigned_count = 0
    for value in identity_by_pipe_id.values():
        if _safe_text(value):
            direct_assigned_count += 1

    save_path = forms.save_file(
        file_ext="xlsx",
        title="Save Pipe Totals Excel",
        default_name=_default_export_name(),
    )
    if not save_path:
        return
    save_path = _ensure_xlsx_path(save_path)

    try:
        _write_excel_xlsx(save_path, rows)
        export_path = save_path
        export_format = "XLSX"
    except Exception as ex:
        logger.warning("Excel export failed. Falling back to CSV. Error: {}".format(ex))
        csv_path = os.path.splitext(save_path)[0] + ".csv"
        _write_csv(csv_path, rows)
        export_path = csv_path
        export_format = "CSV"

    preview_ids = ", ".join([str(x) for x in ordered_ids[:12]]) if ordered_ids else "<none>"

    output.print_md("# Print Pipe Data")
    output.print_md("Selected system types: {}".format(", ".join(sorted(selected_type_names))))
    output.print_md("Pipes collected: {}".format(len(pipe_elements)))
    output.print_md("Direct Identity Mark values found: {}".format(direct_assigned_count))
    output.print_md("System IDs exported: {}".format(len(rows)))
    output.print_md("Assigned System IDs: {}".format(len(assigned_rows)))
    output.print_md("Unassigned bucket present: {}".format("Yes" if unassigned_rows else "No"))
    output.print_md("Ordered ID preview: {}".format(preview_ids))
    output.print_md("Export format: {}".format(export_format))
    output.print_md("Export path: `{}`".format(export_path))

    forms.alert(
        "Pipe data export complete.\n\n"
        "System Types: {}\n"
        "Pipes: {}\n"
        "Direct Identity Values: {}\n"
        "System IDs: {}\n"
        "Assigned IDs: {}\n"
        "Unassigned Bucket: {}\n"
        "File: {}".format(
            len(selected_type_names),
            len(pipe_elements),
            direct_assigned_count,
            len(rows),
            len(assigned_rows),
            "Yes" if unassigned_rows else "No",
            export_path,
        ),
        title=__title__,
    )


if __name__ == "__main__":
    main()






































































