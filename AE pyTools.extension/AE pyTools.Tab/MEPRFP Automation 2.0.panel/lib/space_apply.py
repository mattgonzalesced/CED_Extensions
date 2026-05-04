# -*- coding: utf-8 -*-
"""
Stage 6 — Spaces apply (Revit-side).

Consumes a list of ``SpacePlacementPlan`` from
``space_placement_workflow`` and turns them into Revit family instances
in one transaction. Element_Linker on every placed element carries the
new ``space_id`` + ``space_profile_id`` lineage fields so the Spaces
audit and re-placement tools can find them later.

Family / type resolution and parameter writing reuse the helpers in
``placement.py`` so we stay aligned with the equipment side (tiered
symbol lookup, ``SetValueString`` for unit-bearing params, etc.).
"""

import math

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    ElementId,
    ElementTransformUtils,
    Line,
    XYZ,
)
from Autodesk.Revit.DB.Structure import StructuralType  # noqa: E402

from pyrevit import revit  # noqa: E402

import element_linker as _el  # noqa: E402
import element_linker_io as _el_io  # noqa: E402
import placement as _placement  # noqa: E402  -- reuse private helpers


# ---------------------------------------------------------------------
# Apply
# ---------------------------------------------------------------------

class _ApplyResult(object):

    def __init__(self):
        self.placed = []          # [(plan, element)]
        self.failed = []          # [(plan, status, info)]
        self.warnings = []

    @property
    def n_placed(self):
        return len(self.placed)

    @property
    def n_failed(self):
        return len(self.failed)


def apply_plans(doc, plans, action="Place Space Elements (MEPRFP 2.0)"):
    """Materialise every plan inside one Revit transaction.

    Returns an ``_ApplyResult`` summarising successes and failures so
    the UI can render a per-row status without a second pass.
    """
    result = _ApplyResult()
    if not plans:
        return result

    with revit.Transaction(action, doc=doc):
        for plan in plans:
            try:
                _apply_one(doc, plan, result)
            except Exception as exc:
                result.failed.append((plan, "exception", {"message": str(exc)}))
                result.warnings.append(
                    "Exception placing {!r} in space {}: {}".format(
                        plan.label, plan.space_element_id, exc
                    )
                )

    return result


def _apply_one(doc, plan, result):
    label = plan.label or ""
    family_name, type_name = _placement._split_label(label)
    if not family_name:
        result.failed.append((plan, "no_label", {}))
        return

    symbol, status, available_types = _placement._resolve_family_symbol(
        doc, family_name, type_name, allow_type_substitution=False,
    )
    info = {
        "requested_family": family_name,
        "requested_type": type_name,
        "available_types": available_types,
    }
    if symbol is None:
        result.failed.append((plan, status, info))
        return

    _placement._activate_symbol(symbol)

    target_pt = XYZ(plan.world_pt[0], plan.world_pt[1], plan.world_pt[2])
    level_id = _level_id_for_space(plan)
    level = doc.GetElement(ElementId(int(level_id))) if level_id else None

    try:
        if level is not None:
            inst = doc.Create.NewFamilyInstance(
                target_pt, symbol, level, StructuralType.NonStructural,
            )
        else:
            inst = doc.Create.NewFamilyInstance(
                target_pt, symbol, StructuralType.NonStructural,
            )
    except Exception:
        try:
            inst = doc.Create.NewFamilyInstance(
                target_pt, symbol, StructuralType.NonStructural,
            )
        except Exception as exc:
            result.failed.append(
                (plan, "create_failed", {"message": str(exc), **info})
            )
            return

    if inst is None:
        result.failed.append((plan, "create_failed", info))
        return

    # Rotate about Z if requested.
    if abs(plan.rotation_deg) > 1e-6:
        try:
            axis = Line.CreateBound(
                target_pt,
                XYZ(target_pt.X, target_pt.Y, target_pt.Z + 1.0),
            )
            ElementTransformUtils.RotateElement(
                doc, inst.Id, axis, math.radians(plan.rotation_deg),
            )
        except Exception:
            pass  # non-fatal; element is still placed

    # Write LED-captured parameters.
    led_params = plan.led.parameters if plan.led is not None else None
    if isinstance(led_params, dict) and led_params:
        try:
            _placement._apply_static_parameters(inst, led_params)
        except Exception as exc:
            result.warnings.append(
                "Parameter write failed for ElementId {}: {}".format(
                    inst.Id, exc
                )
            )

    # Stamp Element_Linker with full Spaces lineage.
    try:
        _stamp_linker(inst, plan)
    except Exception as exc:
        result.warnings.append(
            "Element_Linker write failed for ElementId {}: {}".format(
                inst.Id, exc
            )
        )

    result.placed.append((plan, inst))


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------

def _level_id_for_space(plan):
    space = plan.space
    if space is None or space.element is None:
        return None
    try:
        lvl_id = space.element.LevelId
    except Exception:
        return None
    if lvl_id is None:
        return None
    for attr in ("Value", "IntegerValue"):
        try:
            v = getattr(lvl_id, attr)
        except Exception:
            v = None
        if v is None:
            continue
        try:
            return int(v)
        except Exception:
            continue
    return None


def _stamp_linker(elem, plan):
    """Write the Element_Linker with Spaces lineage fields."""
    led = plan.led
    led_id = led.id if led is not None else None
    led_params = led.parameters if (led is not None and isinstance(led.parameters, dict)) else {}

    # Resolved physical state — mirrors what equipment placement records.
    location_ft = None
    rotation_deg = float(plan.rotation_deg or 0.0)
    try:
        loc = elem.Location
        pt = getattr(loc, "Point", None)
        if pt is not None:
            location_ft = [float(pt.X), float(pt.Y), float(pt.Z)]
    except Exception:
        pass

    facing = None
    try:
        f = elem.FacingOrientation
        facing = [float(f.X), float(f.Y), float(f.Z)]
    except Exception:
        facing = None

    payload = _el.ElementLinker(
        led_id=led_id,
        set_id=plan.set_id or None,
        location_ft=location_ft,
        rotation_deg=rotation_deg,
        parent_rotation_deg=None,        # spaces have no parent rotation
        parent_element_id=None,          # the Space is referenced via space_id
        level_id=_element_level_id_value(elem),
        element_id=_element_id_value(elem),
        facing=facing,
        host_name=plan.space.name if plan.space is not None else None,
        parent_location_ft=None,
        ckt_circuit_number=_param_str(led_params, "CKT_Circuit Number_CEDT"),
        ckt_panel=_param_str(led_params, "CKT_Panel_CEDT"),
        # Stage 6 lineage
        space_id=plan.space_element_id,
        space_profile_id=plan.profile_id,
    )
    _el_io.write_to_element(elem, payload)


def _element_id_value(elem):
    if elem is None:
        return None
    eid = getattr(elem, "Id", None)
    if eid is None:
        return None
    for attr in ("Value", "IntegerValue"):
        try:
            v = getattr(eid, attr)
        except Exception:
            v = None
        if v is None:
            continue
        try:
            return int(v)
        except Exception:
            continue
    return None


def _element_level_id_value(elem):
    lid = getattr(elem, "LevelId", None)
    if lid is None:
        return None
    for attr in ("Value", "IntegerValue"):
        try:
            v = getattr(lid, attr)
        except Exception:
            v = None
        if v is None:
            continue
        try:
            return int(v)
        except Exception:
            continue
    return None


def _param_str(params, name):
    if not isinstance(params, dict):
        return None
    value = params.get(name)
    if value is None:
        return None
    if isinstance(value, dict):
        return None
    text = str(value).strip()
    return text or None
