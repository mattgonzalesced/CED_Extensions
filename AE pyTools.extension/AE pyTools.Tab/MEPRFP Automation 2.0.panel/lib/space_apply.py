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
    BuiltInParameter,
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

    # Drain any space_anchored anchor-resolution diagnostics that
    # accumulated during the dry-run collect (the workflow calls
    # ``expand_led_placements`` to compute world_pt for every plan).
    # We pull them BEFORE the apply loop so we capture the values the
    # placement is actually about to use, not anything the apply loop
    # might add.
    try:
        for line in _placement.drain_space_anchored_diagnostics():
            result.warnings.append(line)
    except Exception:
        pass

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
    # Informational plans (e.g. door-relative LED in a doorless
    # space) carry a comment and no world point. Surface them as
    # "no_anchor" failures so the preview row's Status column shows
    # the reason without trying to actually create anything.
    if not getattr(plan, "is_placeable", True) or plan.world_pt is None:
        result.failed.append((plan, "no_anchor", {
            "comment": getattr(plan, "comment", "") or "",
        }))
        return

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

    # ---------------------------------------------------------------
    # Keynote / annotation branch — view-based placement.
    #
    # Keynote symbols (family ``GA_Keynote Symbol_CED``) are 2D
    # annotations bound to a view, not 3D family instances. They MUST
    # be created via the view-aware ``NewFamilyInstance(point, symbol,
    # view)`` overload — the 3-arg overload below would place a
    # phantom instance in model space at the level, which is invisible
    # in plan views and shows up at the wrong Z in 3D. We route any
    # LED flagged as ``is_keynote: true`` (set by the capture engine)
    # or whose family matches the keynote family name through this
    # path, using ``doc.ActiveView`` as the host view.
    # ---------------------------------------------------------------
    is_keynote_led = False
    if plan.led is not None:
        try:
            is_keynote_led = bool(plan.led._data.get("is_keynote"))
        except Exception:
            is_keynote_led = False
    if not is_keynote_led and family_name == "GA_Keynote Symbol_CED":
        is_keynote_led = True

    if is_keynote_led:
        active_view = doc.ActiveView
        # The view-aware overload needs a 2D-ish point on the view's
        # plane. Snap target_pt.Z to the space's level elevation so
        # the keynote lands on the right plan and not at an
        # arbitrary world Z carried from the captured XY.
        if level is not None:
            try:
                level_z = float(level.Elevation or 0.0)
            except Exception:
                level_z = target_pt.Z
            kn_pt = XYZ(target_pt.X, target_pt.Y, level_z)
        else:
            kn_pt = XYZ(target_pt.X, target_pt.Y, 0.0)
        try:
            inst = doc.Create.NewFamilyInstance(kn_pt, symbol, active_view)
        except Exception as exc:
            result.failed.append(
                (plan, "create_failed", {"message": str(exc), **info})
            )
            return
        if inst is None:
            result.failed.append((plan, "create_failed", info))
            return
        # Write captured params (Keynote Value, Keynote Description, etc.).
        led_params = plan.led.parameters if plan.led is not None else None
        if isinstance(led_params, dict) and led_params:
            try:
                _placement._apply_static_parameters(
                    inst, led_params, warnings=result.warnings,
                )
            except Exception as exc:
                result.warnings.append(
                    "Parameter write failed for keynote ElementId {}: {}".format(
                        inst.Id, exc
                    )
                )
        # Stamp the Element_Linker for lineage.
        try:
            _stamp_linker(inst, plan)
        except Exception as exc:
            result.warnings.append(
                "Element_Linker write failed for keynote ElementId {}: {}".format(
                    inst.Id, exc
                )
            )
        result.placed.append((plan, inst))
        return

    # Match the equipment-side ``placement._place_fixture`` byte-for-
    # byte: 3-arg NewFamilyInstance, then rotate, then write the FULL
    # captured params dict (including "Elevation from Level"). The
    # parameter-write path is what actually moves the instance to its
    # captured elevation on level-based families — NewFamilyInstance
    # itself frequently clamps to ``level.Elevation +
    # Default_Elevation``, but ``_set_param_value`` (with the feet-
    # inches fallback parser) lands the captured value via the
    # parameter system and Revit recomputes the geometry from the
    # parameter. This is why equipment placements end up at the right
    # height: not because NewFamilyInstance honoured target_pt.Z, but
    # because the subsequent parameter write does the heavy lifting.
    led_params = plan.led.parameters if plan.led is not None else None

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

    # ---------------------------------------------------------------
    # DIAGNOSTIC: dump everything we know about the elevation pipeline
    # right before _apply_static_parameters runs, then again after.
    # If the parameter still doesn't take, this surfaces exactly where
    # the disconnect is: missing from params dict, write returned
    # False, write returned True but readback is wrong, etc.
    # ---------------------------------------------------------------
    try:
        led_data = plan.led._data if plan.led is not None else {}
        offsets_raw = led_data.get("offsets") or []
        offset0_z = ""
        if offsets_raw and isinstance(offsets_raw[0], dict):
            offset0_z = offsets_raw[0].get("z_inches", "")
        params_raw = led_data.get("parameters") or {}
        elev_in_dict = params_raw.get("Elevation from Level", "<missing>")

        # What's the parameter on the placed instance look like BEFORE
        # we write?
        pre_elev_str = "?"
        pre_elev_double = "?"
        try:
            ep = inst.LookupParameter("Elevation from Level")
            if ep is not None:
                try:
                    pre_elev_str = ep.AsValueString() or "<empty>"
                except Exception:
                    pre_elev_str = "<no AsValueString>"
                try:
                    pre_elev_double = "{:.4f}".format(ep.AsDouble())
                except Exception:
                    pass
                pre_elev_str += " (storage={}, readonly={})".format(
                    getattr(ep.StorageType, "ToString", lambda: "?")(),
                    ep.IsReadOnly,
                )
            else:
                pre_elev_str = "<param not on instance>"
        except Exception as exc:
            pre_elev_str = "<lookup error: {}>".format(exc)

        result.warnings.append(
            "[diag/pre] LED {} | offset[0].z_inches={!r} | "
            "params['Elevation from Level']={!r} | "
            "instance.Elevation from Level (pre-write)={} "
            "AsDouble(pre)={}".format(
                plan.led_id or "?",
                offset0_z, elev_in_dict,
                pre_elev_str, pre_elev_double,
            )
        )
    except Exception:
        pass

    # Write LED-captured parameters straight through, byte-for-byte
    # like the equipment-side ``execute_placement`` does at
    # ``placement.py:_apply_static_parameters(placed,
    # led.get("parameters"))``.
    if isinstance(led_params, dict) and led_params:
        try:
            _placement._apply_static_parameters(
                inst, led_params, warnings=result.warnings,
            )
        except Exception as exc:
            result.warnings.append(
                "Parameter write failed for ElementId {}: {}".format(
                    inst.Id, exc
                )
            )

    # Also try writing the parameter directly here, two ways, with
    # explicit return-value capture, so we can see exactly what each
    # API call returns. If _apply_static_parameters' SetValueString
    # silently fails, this redundant attempt will tell us.
    try:
        z_inches_val = 0.0
        if offsets_raw and isinstance(offsets_raw[0], dict):
            try:
                z_inches_val = float(offsets_raw[0].get("z_inches") or 0.0)
            except Exception:
                z_inches_val = 0.0
        elev_ft = float(z_inches_val) / 12.0
        ep = inst.LookupParameter("Elevation from Level")
        svs_result = "<no param>"
        set_result = "<no param>"
        bip_set_result = "<no param>"
        if ep is not None and not ep.IsReadOnly:
            try:
                feet_int = int(elev_ft)
                inches_part = (elev_ft - feet_int) * 12
                if abs(inches_part - round(inches_part)) < 1e-6:
                    inches_str = '{}"'.format(int(round(inches_part)))
                else:
                    inches_str = '{:g}"'.format(inches_part)
                ftin = "{}' - {}".format(feet_int, inches_str)
                svs_result = repr(ep.SetValueString(ftin))
            except Exception as exc:
                svs_result = "<exception: {}>".format(exc)
            try:
                set_result = repr(ep.Set(float(elev_ft)))
            except Exception as exc:
                set_result = "<exception: {}>".format(exc)
        # Try the BuiltIn slot too
        try:
            bip_param = inst.get_Parameter(
                BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM,
            )
            if bip_param is not None and not bip_param.IsReadOnly:
                bip_set_result = repr(bip_param.Set(float(elev_ft)))
            elif bip_param is None:
                bip_set_result = "<INSTANCE_FREE_HOST_OFFSET_PARAM None>"
            else:
                bip_set_result = "<INSTANCE_FREE_HOST_OFFSET_PARAM read-only>"
        except Exception as exc:
            bip_set_result = "<exception: {}>".format(exc)
        # Read back
        post_elev_str = "?"
        post_elev_double = "?"
        try:
            ep2 = inst.LookupParameter("Elevation from Level")
            if ep2 is not None:
                try:
                    post_elev_str = ep2.AsValueString() or "<empty>"
                except Exception:
                    pass
                try:
                    post_elev_double = "{:.4f}".format(ep2.AsDouble())
                except Exception:
                    pass
        except Exception:
            pass
        result.warnings.append(
            "[diag/post] LED {} | elev_ft={:.4f} | "
            "SetValueString={} | Set(double)={} | "
            "BuiltIn.Set(double)={} | "
            "Elevation from Level (post)={} AsDouble(post)={}".format(
                plan.led_id or "?",
                elev_ft, svs_result, set_result, bip_set_result,
                post_elev_str, post_elev_double,
            )
        )
    except Exception as exc:
        result.warnings.append(
            "[diag/post] LED {} | exception: {}".format(
                plan.led_id or "?", exc,
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
