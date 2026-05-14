# -*- coding: utf-8 -*-
"""
Audit Circuits — read-only health check.

Walks every electrically-relevant placed instance + every existing
``ElectricalSystem`` in the doc and produces a ``CircuitAuditResult``
populated with categorised findings. Mirrors the QAQC pattern:

    A  Missing circuit data        Element has no CKT_Panel_CEDT or
                                   no CKT_Circuit Number_CEDT.
    B  Drift from YAML             Actual CKT data on the live element
                                   doesn't match what the LED's
                                   parameters / Element_Linker expect.
    C  Phantom panel               CKT_Panel_CEDT names a panel that
                                   isn't loaded in the project.
    D  Orphan circuit              ElectricalSystem with zero members
                                   or no panel assignment.
    E  Pole mismatch               Element's Number_of_Poles_CED
                                   exceeds the host system's pole
                                   count, or the system was created
                                   for a 3-pole element but is now
                                   single-pole.

Each finding optionally carries a ``fix_kind`` describing what (if
any) automated remediation can be applied. The audit window dispatches
those fixes.

The workflow is read-only — no transactions, no parameter writes —
so it can be opened modeless without ExternalEvent. Fixes are
applied on demand and route through their own brief transaction.
"""

import math

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
)
from Autodesk.Revit.DB.Electrical import ElectricalSystem  # noqa: E402

import circuit_apply as _apply
import circuit_workflow as _workflow
import element_linker_io as _el_io


# ---------------------------------------------------------------------
# Categories
# ---------------------------------------------------------------------

CAT_A = "A"
CAT_B = "B"
CAT_C = "C"
CAT_D = "D"
CAT_E = "E"

CAT_ALL = (CAT_A, CAT_B, CAT_C, CAT_D, CAT_E)

CAT_LABELS = {
    CAT_A: "A  Missing circuit data",
    CAT_B: "B  Drift from YAML",
    CAT_C: "C  Phantom panel",
    CAT_D: "D  Orphan circuit",
    CAT_E: "E  Pole mismatch",
}


# Fix-kind dispatch.
FIX_NONE = "none"
FIX_RUN_SUPERCIRCUIT = "run_supercircuit"   # cat A — trigger SuperCircuit on this element
FIX_REWRITE_CKT_FROM_YAML = "rewrite_from_yaml"  # cat B — push YAML LED params onto element
FIX_DELETE_ORPHAN = "delete_orphan"          # cat D — delete the empty system
FIX_CLEAR_PANEL_REF = "clear_panel"          # cat C — empty out CKT_Panel_CEDT


# ---------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------

class CircuitAuditFinding(object):
    __slots__ = (
        "category",
        "category_label",
        "element_id",
        "system_id",
        "message",
        "fix_kind",
        "fix_payload",
        # Context fields surfaced on the row so the user can identify
        # which equipment a finding belongs to without having to Select
        # / Zoom first.
        "profile_id",
        "profile_name",
        "family_type_label",
    )

    def __init__(self, category, element_id=None, system_id=None,
                 message="", fix_kind=FIX_NONE, fix_payload=None,
                 profile_id=None, profile_name=None, family_type_label=None):
        self.category = category
        self.category_label = CAT_LABELS.get(category, category)
        self.element_id = element_id
        self.system_id = system_id
        self.message = message
        self.fix_kind = fix_kind
        self.fix_payload = dict(fix_payload or {})
        self.profile_id = profile_id or None
        self.profile_name = profile_name or None
        self.family_type_label = family_type_label or None


class CircuitAuditResult(object):
    def __init__(self):
        self.findings = []
        self.counts = {c: 0 for c in CAT_ALL}

    def add(self, finding):
        self.findings.append(finding)
        self.counts[finding.category] = self.counts.get(finding.category, 0) + 1


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------

def _id_value(elem_or_id):
    if elem_or_id is None:
        return None
    eid = getattr(elem_or_id, "Id", None) or elem_or_id
    return getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)


def _read_param_string(elem, name):
    if elem is None:
        return ""
    try:
        p = elem.LookupParameter(name)
    except Exception:
        return ""
    if p is None:
        return ""
    try:
        s = p.AsString()
        return s or ""
    except Exception:
        try:
            return p.AsValueString() or ""
        except Exception:
            return ""


def _read_param_int(elem, name):
    if elem is None:
        return None
    try:
        p = elem.LookupParameter(name)
    except Exception:
        return None
    if p is None:
        return None
    try:
        return int(p.AsInteger())
    except Exception:
        try:
            return int(float(p.AsString() or ""))
        except Exception:
            return None


def _yaml_expected_for_led(led):
    """Pull the captured CKT data from the LED's parameters dict.

    Values are coerced through ``str(...)`` before stripping because
    the LED's params dict can carry non-string types — circuit numbers
    are sometimes stored as ints, and parent / sibling directives are
    stored as dicts. ``int.strip()`` / ``dict.strip()`` would crash
    refresh; the directive case is intentionally rendered as its repr
    here (audit time can't resolve them — that's the placement
    engine's job — so leaving them as text is the cleanest "no
    expected value" outcome).
    """
    if not isinstance(led, dict):
        return {}
    params = led.get("parameters") or {}
    if not isinstance(params, dict):
        return {}

    def _as_text(value):
        if value is None:
            return ""
        if isinstance(value, dict):
            # Directive (BYPARENT/BYSIBLING) — leave empty so the diff
            # below treats it as "no expected value" and the row
            # doesn't drift-flag against the live element.
            return ""
        return str(value).strip()

    return {
        "panel": _as_text(params.get("CKT_Panel_CEDT")),
        "circuit": _as_text(params.get("CKT_Circuit Number_CEDT")),
        "load": _as_text(params.get("CKT_Load Name_CEDT")),
    }


def _element_family_type_label(elem):
    """``"Family : Type"`` for a placed FamilyInstance / Group, or ``""``
    when the symbol can't be resolved. Surfaced on findings so the user
    can identify the equipment without selecting it first."""
    if elem is None:
        return ""
    try:
        sym = getattr(elem, "Symbol", None)
        if sym is not None:
            family = getattr(sym, "Family", None)
            family_name = getattr(family, "Name", "") if family is not None else ""
            type_name = getattr(sym, "Name", "") or ""
            if family_name and type_name:
                return "{} : {}".format(family_name, type_name)
            if family_name:
                return family_name
            if type_name:
                return type_name
    except Exception:
        pass
    try:
        gtype = getattr(elem, "GroupType", None)
        if gtype is not None:
            name = getattr(gtype, "Name", "") or ""
            if name:
                return name
    except Exception:
        pass
    return ""


def _profile_context_for_elem(elem, led_index):
    """Return ``(profile_id, profile_name, family_type_label)`` for a
    placed element. Looks up the element's Element_Linker, resolves
    its ``led_id`` against the profile data, and falls back to empty
    strings when any step fails (e.g. no Element_Linker stamped, LED
    deleted from YAML). The caller stamps these onto the finding so
    every row carries equipment-identifying context.
    """
    family_type_label = _element_family_type_label(elem)
    profile_id = ""
    profile_name = ""
    try:
        linker = _el_io.read_from_element(elem)
    except Exception:
        linker = None
    if linker is not None and linker.led_id:
        entry = led_index.get(linker.led_id)
        if entry is not None:
            profile, _set_dict, _led = entry
            profile_id = profile.get("id") or ""
            profile_name = profile.get("name") or ""
    return profile_id, profile_name, family_type_label


def _build_led_index(profile_data):
    out = {}
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
# Audit
# ---------------------------------------------------------------------

def run_audit(doc, profile_data=None, categories=None):
    """Walk the doc and emit a populated ``CircuitAuditResult``.

    ``categories`` restricts the run to a subset (set/list of CAT_*);
    None means all five.
    """
    requested = set(categories) if categories else set(CAT_ALL)
    result = CircuitAuditResult()

    panel_index = _apply.collect_panel_index(doc) if (CAT_A in requested or
                                                      CAT_C in requested or
                                                      CAT_B in requested or
                                                      CAT_E in requested) else {}
    led_index = _build_led_index(profile_data) if profile_data else {}

    # 1. Walk placed instances for cats A / B / C / E.
    if requested & {CAT_A, CAT_B, CAT_C, CAT_E}:
        for elem in _enumerate_target_instances(doc):
            elem_id = _id_value(elem)
            panel_str = _read_param_string(elem, "CKT_Panel_CEDT")
            circuit_str = _read_param_string(elem, "CKT_Circuit Number_CEDT")
            load_str = _read_param_string(elem, "CKT_Load Name_CEDT")
            poles_elem = _read_param_int(elem, "Number of Poles_CED")

            # Resolve per-element context once so every finding from
            # this element carries the same profile / family:type tag.
            ctx_pid, ctx_pname, ctx_label = _profile_context_for_elem(
                elem, led_index,
            )

            # Cat A — missing
            if CAT_A in requested:
                if not panel_str:
                    result.add(CircuitAuditFinding(
                        category=CAT_A,
                        element_id=elem_id,
                        message="CKT_Panel_CEDT is empty",
                        fix_kind=FIX_RUN_SUPERCIRCUIT,
                        profile_id=ctx_pid,
                        profile_name=ctx_pname,
                        family_type_label=ctx_label,
                    ))
                if not circuit_str:
                    result.add(CircuitAuditFinding(
                        category=CAT_A,
                        element_id=elem_id,
                        message="CKT_Circuit Number_CEDT is empty",
                        fix_kind=FIX_RUN_SUPERCIRCUIT,
                        profile_id=ctx_pid,
                        profile_name=ctx_pname,
                        family_type_label=ctx_label,
                    ))

            # Cat C — phantom panel
            if CAT_C in requested and panel_str:
                # First token only — multi-panel strings count as
                # phantom only when no token resolves.
                tokens = [t.strip() for t in panel_str.replace(",", " ").replace(";", " ")
                          .replace("|", " ").replace("/", " ").split()]
                tokens = [t for t in tokens if t]
                resolved = any(t.lower() in panel_index for t in tokens)
                if not resolved:
                    result.add(CircuitAuditFinding(
                        category=CAT_C,
                        element_id=elem_id,
                        message="Panel '{}' is not loaded in the project".format(panel_str),
                        fix_kind=FIX_CLEAR_PANEL_REF,
                        profile_id=ctx_pid,
                        profile_name=ctx_pname,
                        family_type_label=ctx_label,
                    ))

            # Cat B — drift from YAML
            if CAT_B in requested:
                linker = None
                try:
                    linker = _el_io.read_from_element(elem)
                except Exception:
                    linker = None
                if linker is not None and linker.led_id:
                    entry = led_index.get(linker.led_id)
                    if entry is not None:
                        _, _, led = entry
                        expected = _yaml_expected_for_led(led)
                        drift = _diff_ckt(expected, panel_str, circuit_str, load_str)
                        if drift:
                            result.add(CircuitAuditFinding(
                                category=CAT_B,
                                element_id=elem_id,
                                message="Drift from YAML: {}".format("; ".join(drift)),
                                fix_kind=FIX_REWRITE_CKT_FROM_YAML,
                                fix_payload={"led_id": linker.led_id},
                                profile_id=ctx_pid,
                                profile_name=ctx_pname,
                                family_type_label=ctx_label,
                            ))

            # Cat E — pole mismatch (deferred to the system-walk pass)

    # 2. Walk ElectricalSystems for cats D / E.
    if requested & {CAT_D, CAT_E}:
        for system in _enumerate_systems(doc):
            sid = _id_value(system)
            members = list(getattr(system, "Elements", []) or [])
            if CAT_D in requested:
                if not members:
                    result.add(CircuitAuditFinding(
                        category=CAT_D,
                        system_id=sid,
                        message="Circuit has no members",
                        fix_kind=FIX_DELETE_ORPHAN,
                    ))
                else:
                    panel = getattr(system, "BaseEquipment", None)
                    if panel is None:
                        result.add(CircuitAuditFinding(
                            category=CAT_D,
                            system_id=sid,
                            message="Circuit has no panel assignment",
                            fix_kind=FIX_NONE,
                        ))
            if CAT_E in requested and members:
                sys_poles = _system_poles(system)
                for m in members:
                    if m is None:
                        continue
                    member_poles = _read_param_int(m, "Number of Poles_CED")
                    if member_poles and sys_poles and member_poles > sys_poles:
                        ctx_pid, ctx_pname, ctx_label = _profile_context_for_elem(
                            m, led_index,
                        )
                        result.add(CircuitAuditFinding(
                            category=CAT_E,
                            element_id=_id_value(m),
                            system_id=sid,
                            message=(
                                "Element wants {} pole(s) but circuit is {} pole(s)"
                            ).format(member_poles, sys_poles),
                            fix_kind=FIX_NONE,
                            profile_id=ctx_pid,
                            profile_name=ctx_pname,
                            family_type_label=ctx_label,
                        ))

    return result


# ---------------------------------------------------------------------
# Walkers
# ---------------------------------------------------------------------

def _enumerate_target_instances(doc):
    target_categories = (
        BuiltInCategory.OST_ElectricalFixtures,
        BuiltInCategory.OST_DataDevices,
        BuiltInCategory.OST_CommunicationDevices,
        BuiltInCategory.OST_FireAlarmDevices,
        BuiltInCategory.OST_NurseCallDevices,
        BuiltInCategory.OST_SecurityDevices,
        BuiltInCategory.OST_TelephoneDevices,
        BuiltInCategory.OST_MechanicalControlDevices,
    )
    for cat in target_categories:
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
            yield elem


def _enumerate_systems(doc):
    try:
        collector = (
            FilteredElementCollector(doc)
            .OfClass(ElectricalSystem)
            .WhereElementIsNotElementType()
        )
    except Exception:
        return
    for s in collector:
        yield s


def _system_poles(system):
    try:
        param = system.get_Parameter(BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)
        if param is not None:
            return int(param.AsInteger() or 0) or None
    except Exception:
        return None
    return None


def _diff_ckt(expected, actual_panel, actual_circuit, actual_load):
    """Return a list of diff descriptions where YAML expected differs
    from actual element parameters. Empty list = no drift."""
    def _norm(value):
        if value is None:
            return ""
        return str(value).strip().lower()

    out = []
    exp_panel = _norm(expected.get("panel"))
    exp_circuit = _norm(expected.get("circuit"))
    exp_load = _norm(expected.get("load"))
    if exp_panel and exp_panel != _norm(actual_panel):
        out.append("panel '{}' vs YAML '{}'".format(actual_panel, expected["panel"]))
    if exp_circuit and exp_circuit != _norm(actual_circuit):
        out.append("circuit '{}' vs YAML '{}'".format(actual_circuit, expected["circuit"]))
    if exp_load and exp_load != _norm(actual_load):
        out.append("load '{}' vs YAML '{}'".format(actual_load, expected["load"]))
    return out


# ---------------------------------------------------------------------
# Fixes (read-only callers wrap in their own transaction)
# ---------------------------------------------------------------------

def execute_fix(doc, profile_data, finding):
    """Apply the auto-fix for one finding. Caller manages the
    transaction. Returns ``(ok: bool, message: str)``.
    """
    if finding.fix_kind == FIX_NONE:
        return False, "No automated fix for this category."

    if finding.fix_kind == FIX_DELETE_ORPHAN:
        sid = finding.system_id
        if sid is None:
            return False, "No system id."
        try:
            doc.Delete(ElementId(int(sid)))
        except Exception as exc:
            return False, "Delete failed: {}".format(exc)
        return True, "Orphan circuit deleted."

    if finding.fix_kind == FIX_CLEAR_PANEL_REF:
        elem = _resolve_element(doc, finding.element_id)
        if elem is None:
            return False, "Element not found."
        try:
            p = elem.LookupParameter("CKT_Panel_CEDT")
            if p is not None and not p.IsReadOnly:
                p.Set("")
                return True, "Phantom CKT_Panel_CEDT cleared."
        except Exception as exc:
            return False, "Clear failed: {}".format(exc)
        return False, "CKT_Panel_CEDT not writable."

    if finding.fix_kind == FIX_REWRITE_CKT_FROM_YAML:
        elem = _resolve_element(doc, finding.element_id)
        if elem is None:
            return False, "Element not found."
        led_id = (finding.fix_payload or {}).get("led_id")
        if not led_id:
            return False, "No led_id on finding."
        led_index = _build_led_index(profile_data)
        entry = led_index.get(led_id)
        if entry is None:
            return False, "LED {} not in active YAML store.".format(led_id)
        _, _, led = entry
        expected = _yaml_expected_for_led(led)
        wrote = []
        for revit_name, key in (
            ("CKT_Panel_CEDT", "panel"),
            ("CKT_Circuit Number_CEDT", "circuit"),
            ("CKT_Load Name_CEDT", "load"),
        ):
            value = expected.get(key) or ""
            if not value:
                continue
            try:
                p = elem.LookupParameter(revit_name)
                if p is not None and not p.IsReadOnly:
                    p.Set(value)
                    wrote.append(revit_name)
            except Exception:
                continue
        if not wrote:
            return False, "No CKT_* parameters were writable on the element."
        return True, "Rewrote {} from YAML.".format(", ".join(wrote))

    if finding.fix_kind == FIX_RUN_SUPERCIRCUIT:
        return False, ("No automated fix for missing CKT data — "
                       "run SuperCircuit V5 on this element to assign one.")

    return False, "Unknown fix_kind: {}".format(finding.fix_kind)


def _resolve_element(doc, element_id):
    if element_id is None:
        return None
    try:
        return doc.GetElement(ElementId(int(element_id)))
    except Exception:
        return None
