# -*- coding: utf-8 -*-
"""
Proximity check for lighting fixtures near CED-R-KRACK coils.
Runs after sync when enabled.
"""

import math
import os
import sys

from pyrevit import DB, forms, revit, script

LIB_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "CEDLib.lib")
)
if LIB_ROOT not in sys.path:
    sys.path.append(LIB_ROOT)

from ExtensibleStorage import ExtensibleStorage  # noqa: E402

SETTING_KEY = "proximity_lights_coils_check"
FAMILY_PREFIX = "CED-R-KRACK"
THRESHOLD_INCHES = 18.0
THRESHOLD_FEET = THRESHOLD_INCHES / 12.0
GEOM_DETAIL_LEVEL = DB.ViewDetailLevel.Fine


def _get_doc(doc=None):
    if doc is not None:
        return doc
    try:
        return getattr(revit, "doc", None)
    except Exception:
        return None


def get_setting(default=True, doc=None):
    doc = _get_doc(doc)
    if doc is None:
        return bool(default)
    value = ExtensibleStorage.get_user_setting(doc, SETTING_KEY, default=None)
    if value is None:
        return bool(default)
    return bool(value)


def set_setting(value, doc=None):
    doc = _get_doc(doc)
    if doc is None:
        return False
    return ExtensibleStorage.set_user_setting(doc, SETTING_KEY, bool(value))


def _family_type_label(elem):
    if elem is None:
        return "<missing>"
    label = None
    try:
        symbol = getattr(elem, "Symbol", None)
        family = getattr(symbol, "Family", None) if symbol else None
        fam_name = getattr(family, "Name", None) if family else None
        type_name = getattr(symbol, "Name", None) if symbol else None
        if fam_name and type_name:
            label = "{} : {}".format(fam_name, type_name)
    except Exception:
        label = None
    if not label:
        try:
            label = getattr(elem, "Name", None)
        except Exception:
            label = None
    return label or "<element>"


def _xyz_tuple(xyz):
    return (xyz.X, xyz.Y, xyz.Z)


def _v_sub(a, b):
    return (a[0] - b[0], a[1] - b[1], a[2] - b[2])


def _v_add(a, b):
    return (a[0] + b[0], a[1] + b[1], a[2] + b[2])


def _v_mul(a, s):
    return (a[0] * s, a[1] * s, a[2] * s)


def _v_dot(a, b):
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def _v_len_sq(a):
    return _v_dot(a, a)


def _distance_sq(a, b):
    return _v_len_sq(_v_sub(a, b))


def _segment_distance_sq(p1, q1, p2, q2):
    # From "Real-Time Collision Detection" (Christer Ericson)
    eps = 1e-9
    u = _v_sub(q1, p1)
    v = _v_sub(q2, p2)
    w = _v_sub(p1, p2)
    a = _v_dot(u, u)
    b = _v_dot(u, v)
    c = _v_dot(v, v)
    d = _v_dot(u, w)
    e = _v_dot(v, w)
    denom = a * c - b * b
    sc = 0.0
    sN = 0.0
    sD = denom
    tc = 0.0
    tN = 0.0
    tD = denom

    if denom < eps:
        sN = 0.0
        sD = 1.0
        tN = e
        tD = c
    else:
        sN = (b * e - c * d)
        tN = (a * e - b * d)
        if sN < 0.0:
            sN = 0.0
            tN = e
            tD = c
        elif sN > sD:
            sN = sD
            tN = e + b
            tD = c

    if tN < 0.0:
        tN = 0.0
        if -d < 0.0:
            sN = 0.0
        elif -d > a:
            sN = sD
        else:
            sN = -d
            sD = a
    elif tN > tD:
        tN = tD
        if (-d + b) < 0.0:
            sN = 0.0
        elif (-d + b) > a:
            sN = sD
        else:
            sN = (-d + b)
            sD = a

    sc = 0.0 if abs(sN) < eps else sN / sD
    tc = 0.0 if abs(tN) < eps else tN / tD
    dP = _v_sub(_v_add(w, _v_mul(u, sc)), _v_mul(v, tc))
    return _v_len_sq(dP)


def _point_triangle_distance_sq(p, a, b, c):
    # From "Real-Time Collision Detection" (Christer Ericson)
    ab = _v_sub(b, a)
    ac = _v_sub(c, a)
    ap = _v_sub(p, a)
    d1 = _v_dot(ab, ap)
    d2 = _v_dot(ac, ap)
    if d1 <= 0.0 and d2 <= 0.0:
        return _distance_sq(p, a)

    bp = _v_sub(p, b)
    d3 = _v_dot(ab, bp)
    d4 = _v_dot(ac, bp)
    if d3 >= 0.0 and d4 <= d3:
        return _distance_sq(p, b)

    vc = d1 * d4 - d3 * d2
    if vc <= 0.0 and d1 >= 0.0 and d3 <= 0.0:
        v = d1 / (d1 - d3)
        proj = _v_add(a, _v_mul(ab, v))
        return _distance_sq(p, proj)

    cp = _v_sub(p, c)
    d5 = _v_dot(ab, cp)
    d6 = _v_dot(ac, cp)
    if d6 >= 0.0 and d5 <= d6:
        return _distance_sq(p, c)

    vb = d5 * d2 - d1 * d6
    if vb <= 0.0 and d2 >= 0.0 and d6 <= 0.0:
        w = d2 / (d2 - d6)
        proj = _v_add(a, _v_mul(ac, w))
        return _distance_sq(p, proj)

    va = d3 * d6 - d5 * d4
    if va <= 0.0 and (d4 - d3) >= 0.0 and (d5 - d6) >= 0.0:
        w = (d4 - d3) / ((d4 - d3) + (d5 - d6))
        proj = _v_add(b, _v_mul(_v_sub(c, b), w))
        return _distance_sq(p, proj)

    # Inside face region
    denom = 1.0 / (va + vb + vc)
    v = vb * denom
    w = vc * denom
    proj = _v_add(a, _v_add(_v_mul(ab, v), _v_mul(ac, w)))
    return _distance_sq(p, proj)


def _triangle_distance_sq(tri_a, tri_b):
    a0, a1, a2 = tri_a
    b0, b1, b2 = tri_b
    min_sq = None
    for p in (a0, a1, a2):
        d = _point_triangle_distance_sq(p, b0, b1, b2)
        min_sq = d if min_sq is None else min(min_sq, d)
    for p in (b0, b1, b2):
        d = _point_triangle_distance_sq(p, a0, a1, a2)
        min_sq = d if min_sq is None else min(min_sq, d)
    edges_a = ((a0, a1), (a1, a2), (a2, a0))
    edges_b = ((b0, b1), (b1, b2), (b2, b0))
    for ea in edges_a:
        for eb in edges_b:
            d = _segment_distance_sq(ea[0], ea[1], eb[0], eb[1])
            min_sq = d if min_sq is None else min(min_sq, d)
    return min_sq or 0.0


def _triangles_distance_sq(tris_a, tris_b):
    min_sq = None
    for tri_a in tris_a:
        for tri_b in tris_b:
            d = _triangle_distance_sq(tri_a, tri_b)
            min_sq = d if min_sq is None else min(min_sq, d)
            if min_sq == 0.0:
                return 0.0
    return min_sq if min_sq is not None else None


def _bbox_distance(bbox_a, bbox_b):
    if bbox_a is None or bbox_b is None:
        return None

    def axis_distance(a_min, a_max, b_min, b_max):
        if a_max < b_min:
            return b_min - a_max
        if b_max < a_min:
            return a_min - b_max
        return 0.0

    dx = axis_distance(bbox_a.Min.X, bbox_a.Max.X, bbox_b.Min.X, bbox_b.Max.X)
    dy = axis_distance(bbox_a.Min.Y, bbox_a.Max.Y, bbox_b.Min.Y, bbox_b.Max.Y)
    dz = axis_distance(bbox_a.Min.Z, bbox_a.Max.Z, bbox_b.Min.Z, bbox_b.Max.Z)
    return math.sqrt(dx * dx + dy * dy + dz * dz)


def _collect_triangles_from_geom(geom, triangles):
    if geom is None:
        return
    for obj in geom:
        if isinstance(obj, DB.Solid):
            try:
                if obj.Volume <= 1e-9:
                    continue
            except Exception:
                pass
            try:
                faces = obj.Faces
            except Exception:
                faces = None
            if faces is None:
                continue
            for face in faces:
                try:
                    mesh = face.Triangulate()
                except Exception:
                    mesh = None
                if mesh is None:
                    continue
                for i in range(mesh.NumTriangles):
                    tri = mesh.get_Triangle(i)
                    triangles.append((
                        _xyz_tuple(tri.get_Vertex(0)),
                        _xyz_tuple(tri.get_Vertex(1)),
                        _xyz_tuple(tri.get_Vertex(2)),
                    ))
        elif isinstance(obj, DB.Mesh):
            for i in range(obj.NumTriangles):
                tri = obj.get_Triangle(i)
                triangles.append((
                    _xyz_tuple(tri.get_Vertex(0)),
                    _xyz_tuple(tri.get_Vertex(1)),
                    _xyz_tuple(tri.get_Vertex(2)),
                ))
        elif isinstance(obj, DB.GeometryInstance):
            try:
                inst_geom = obj.GetInstanceGeometry()
            except Exception:
                inst_geom = None
            if inst_geom is not None:
                _collect_triangles_from_geom(inst_geom, triangles)
        elif isinstance(obj, DB.GeometryElement):
            _collect_triangles_from_geom(obj, triangles)


def _element_triangles(elem):
    if elem is None:
        return []
    try:
        options = DB.Options()
        options.DetailLevel = GEOM_DETAIL_LEVEL
        options.ComputeReferences = False
        options.IncludeNonVisibleObjects = True
        geom = elem.get_Geometry(options)
    except Exception:
        geom = None
    triangles = []
    _collect_triangles_from_geom(geom, triangles)
    return triangles


def _collect_coils(doc, option_filter):
    coils = []
    try:
        collector = (
            DB.FilteredElementCollector(doc)
            .OfClass(DB.FamilyInstance)
            .WhereElementIsNotElementType()
            .WherePasses(option_filter)
        )
    except Exception:
        return coils
    prefix = FAMILY_PREFIX.upper()
    for elem in collector:
        try:
            symbol = getattr(elem, "Symbol", None)
            family = getattr(symbol, "Family", None) if symbol else None
            fam_name = getattr(family, "Name", None) if family else None
        except Exception:
            fam_name = None
        if not fam_name:
            continue
        if not fam_name.upper().startswith(prefix):
            continue
        bbox = None
        try:
            bbox = elem.get_BoundingBox(None)
        except Exception:
            bbox = None
        if bbox is None:
            continue
        coils.append({
            "id": elem.Id,
            "label": _family_type_label(elem),
            "bbox": bbox,
            "triangles": _element_triangles(elem),
        })
    return coils


def _collect_lights(doc, option_filter):
    lights = []
    try:
        collector = (
            DB.FilteredElementCollector(doc)
            .OfCategory(DB.BuiltInCategory.OST_LightingFixtures)
            .WhereElementIsNotElementType()
            .WherePasses(option_filter)
        )
    except Exception:
        return lights
    for elem in collector:
        bbox = None
        try:
            bbox = elem.get_BoundingBox(None)
        except Exception:
            bbox = None
        if bbox is None:
            continue
        lights.append({
            "id": elem.Id,
            "label": _family_type_label(elem),
            "bbox": bbox,
            "triangles": _element_triangles(elem),
        })
    return lights


def collect_hits(doc):
    if doc is None:
        return []
    option_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)
    coils = _collect_coils(doc, option_filter)
    lights = _collect_lights(doc, option_filter)
    if not coils or not lights:
        return []
    hits = []
    for light in lights:
        light_bbox = light.get("bbox")
        if light_bbox is None:
            continue
        for coil in coils:
            dist_ft = _bbox_distance(light_bbox, coil.get("bbox"))
            if dist_ft is None:
                continue
            if dist_ft <= THRESHOLD_FEET:
                tris_a = light.get("triangles") or []
                tris_b = coil.get("triangles") or []
                if tris_a and tris_b:
                    dist_sq = _triangles_distance_sq(tris_a, tris_b)
                    if dist_sq is None:
                        continue
                    dist_ft = math.sqrt(dist_sq)
                hits.append({
                    "light_id": light.get("id"),
                    "light_label": light.get("label"),
                    "coil_id": coil.get("id"),
                    "coil_label": coil.get("label"),
                    "distance_ft": dist_ft,
                })
    return hits


def _report_results(results, show_empty=False):
    output = script.get_output()
    output.set_width(1000)
    if not results:
        if show_empty:
            output.print_md("# Proximity Check: Lights-Coils")
            output.print_md("### No issues found.")
            forms.alert("No lighting fixtures found within 18 inches of CED-R-KRACK coils.")
        return
    output.print_md("# Proximity Check: Lights-Coils")
    table = []
    for item in results:
        dist_in = (item.get("distance_ft") or 0.0) * 12.0
        table.append([
            output.linkify(item.get("light_id")),
            item.get("light_label") or "<unknown>",
            output.linkify(item.get("coil_id")),
            item.get("coil_label") or "<unknown>",
            "{:.2f}".format(dist_in),
        ])
    output.print_table(
        table,
        columns=[
            "Lighting ID",
            "Lighting Family : Type",
            "Coil ID",
            "Coil Family : Type",
            "Distance (in)",
        ],
    )
    forms.alert(
        "Found {} lighting fixture(s) within 18 inches of CED-R-KRACK coils.\n\nSee the output panel for details.".format(
            len(results)
        ),
        title="Proximity Check: Lights-Coils",
    )


def run_check(doc=None, show_ui=True, show_empty=False):
    doc = _get_doc(doc)
    if doc is None or getattr(doc, "IsFamilyDocument", False):
        return []
    results = collect_hits(doc)
    if show_ui:
        _report_results(results, show_empty=show_empty)
    return results


def run_sync_check(doc):
    doc = _get_doc(doc)
    if doc is None or getattr(doc, "IsFamilyDocument", False):
        return
    if not get_setting(default=True, doc=doc):
        return
    run_check(doc, show_ui=True, show_empty=False)


__all__ = [
    "get_setting",
    "set_setting",
    "collect_hits",
    "run_check",
    "run_sync_check",
]
