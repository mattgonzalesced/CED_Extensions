# -*- coding: utf-8 -*-
"""
Revit-API edge for SuperCircuit.

Two responsibilities:

  1. **Synchronous primitives** that the workflow uses regardless of
     UI mode — building Revit collections of element ids, looking up
     panels by name, calling ``ElectricalSystem.Create`` + ``SelectPanel``,
     and stamping ``CKT_*`` parameters on the new circuit / its
     members.

  2. **Modeless ExternalEvent gateway** so a non-modal preview window
     can request "apply this batch of groups" and have the actual
     Revit-API mutations run on Revit's main thread inside a single
     transaction. The preview controller calls
     ``request_apply(doc, groups, callback)`` and gets a callback with
     a ``CircuitApplyResult`` once Revit has finished.

The grouping / phasing layers stay pure-Python; everything that
touches the API lives here.
"""

import math

import clr  # noqa: F401  -- needed before importing Autodesk.Revit.DB

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    LocationPoint,
    XYZ,
)
from Autodesk.Revit.DB.Electrical import (  # noqa: E402
    ElectricalSystem,
    ElectricalSystemType,
)
from Autodesk.Revit.UI import (  # noqa: E402
    ExternalEvent,
    IExternalEventHandler,
)
from System.Collections.Generic import List as _NetList  # noqa: E402

import circuit_phasing as _phasing


# ---------------------------------------------------------------------
# Result types
# ---------------------------------------------------------------------

class CircuitApplyResult(object):
    """Outcome of one apply pass."""

    __slots__ = (
        "created_count",
        "skipped_count",
        "failed_count",
        "system_ids_by_group_key",   # {group.key: ElementId.Value (int)}
        "warnings",
    )

    def __init__(self):
        self.created_count = 0
        self.skipped_count = 0
        self.failed_count = 0
        self.system_ids_by_group_key = {}
        self.warnings = []


# ---------------------------------------------------------------------
# Panel + element lookup helpers (synchronous)
# ---------------------------------------------------------------------

def collect_panel_index(doc):
    """Return ``{panel_name_lower: ElectricalEquipment_element}`` for
    every electrical panel in the doc.

    Used by the workflow to resolve ``CKT_Panel_CEDT`` strings before
    calling ``ElectricalSystem.SelectPanel``. Returns the underlying
    Revit element so callers can read its location and supported
    distribution-system list as needed.
    """
    out = {}
    try:
        collector = (
            FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
            .OfClass(FamilyInstance)
            .WhereElementIsNotElementType()
        )
    except Exception:
        return out
    for elem in collector:
        name = ""
        try:
            param = elem.get_Parameter(BuiltInParameter.RBS_ELEC_PANEL_NAME)
            if param is not None:
                name = param.AsString() or ""
        except Exception:
            name = ""
        if not name:
            try:
                name = (elem.Name or "")
            except Exception:
                name = ""
        key = (name or "").strip().lower()
        if not key:
            continue
        # First-seen wins to avoid duplicate-name surprises.
        out.setdefault(key, elem)
    return out


def panel_world_pt(panel_elem):
    """``(x, y, z)`` for a panel's LocationPoint, or None if not point-based."""
    if panel_elem is None:
        return None
    loc = getattr(panel_elem, "Location", None)
    if not isinstance(loc, LocationPoint):
        return None
    try:
        p = loc.Point
        if p is None:
            return None
        return (p.X, p.Y, p.Z)
    except Exception:
        return None


def panel_distribution_system_ids(panel_elem):
    """Return the list of distribution-system ElementId values the
    panel will accept. Implementation walks the panel's ``MEPModel``
    when available, falling back to the panel's
    ``DISTRIBUTION_SYSTEM_PARAM`` for radial-style equipment.
    """
    out = []
    if panel_elem is None:
        return out
    seen = set()

    def _push(eid):
        v = getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)
        if v is None or v in seen or v <= 0:
            return
        seen.add(v)
        out.append(v)

    try:
        param = panel_elem.get_Parameter(BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM)
        if param is not None:
            _push(param.AsElementId())
    except Exception:
        pass

    try:
        mep = getattr(panel_elem, "MEPModel", None)
        if mep is not None:
            for sys_obj in (getattr(mep, "GetAssignedElectricalSystems", None) or (lambda: []))():
                _push(getattr(sys_obj, "Id", None))
    except Exception:
        pass

    return out


# ---------------------------------------------------------------------
# Per-group apply (synchronous — caller manages the transaction)
# ---------------------------------------------------------------------

def _id_value(elem):
    if elem is None:
        return None
    eid = elem.Id
    return getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)


def _to_element_id(value):
    try:
        return ElementId(int(value))
    except Exception:
        return None


def _ensure_panel_element(doc, group, panel_index):
    """Resolve and cache ``group.panel_element`` from the panel index
    using ``group.panel_name``. Returns the element or None."""
    if group.panel_element is not None:
        return group.panel_element
    key = (group.panel_name or "").strip().lower()
    if not key:
        return None
    panel = panel_index.get(key)
    if panel is not None:
        group.panel_element = panel
    return panel


def apply_group(doc, group, panel_index, phase_tracker=None, logger=None):
    """Create one Revit ``ElectricalSystem`` for ``group``.

    Caller must have an open Revit transaction. Returns
    ``(system_or_None, warning_or_None)``. The system, when created, is
    panel-bound and has ``CKT_Load Name_CEDT`` /
    ``CKT_Circuit Number_CEDT`` / ``CKT_Panel_CEDT`` written on each
    member if those parameters exist on the family.
    """
    if group is None or not group.members:
        return None, "Empty group"

    # Members → ElementId list.
    eids = _NetList[ElementId]()
    for member in group.members:
        eid = _to_element_id(member.element_id)
        if eid is not None:
            eids.Add(eid)
    if eids.Count == 0:
        return None, "Group has no resolvable element ids"

    panel_elem = _ensure_panel_element(doc, group, panel_index)

    try:
        system = ElectricalSystem.Create(doc, eids, ElectricalSystemType.PowerCircuit)
    except Exception as exc:
        return None, "ElectricalSystem.Create failed: {}".format(exc)
    if system is None:
        return None, "ElectricalSystem.Create returned None"

    # Bind panel.
    if panel_elem is not None:
        try:
            system.SelectPanel(panel_elem)
        except Exception as exc:
            return system, "SelectPanel failed: {}".format(exc)

    # Stamp the load name on the new system.
    if group.load_name:
        try:
            param = system.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_NAME)
            if param is not None and not param.IsReadOnly:
                param.Set(str(group.load_name))
        except Exception:
            pass

    # Stamp the breaker rating on the system. ``CKT_Rating_CED`` is
    # captured per-fixture; group.rating is the max across members so
    # the breaker covers the heaviest member. SetValueString handles
    # display-unit conversion (so ``20`` lands as ``20 A``).
    if group.rating is not None:
        _set_circuit_param(
            system,
            BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM,
            group.rating,
        )

    # Stamp the circuit's schedule notes from CKT_Schedule Notes_CEDT.
    if group.schedule_notes:
        _set_circuit_param(
            system,
            BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM,
            group.schedule_notes,
        )

    # Mirror the user's CKT_* fields onto the placed members so audits
    # can read them off the elements themselves (matches legacy
    # behaviour). Soft-fail per parameter — missing params are skipped.
    _stamp_member_ckt_fields(group)

    return system, None


def _set_circuit_param(system, builtin, value):
    """Best-effort write of one BuiltInParameter on an ``ElectricalSystem``.

    Routes through ``SetValueString`` for unit-bearing values
    (rating, length, etc.) and falls back to ``Set`` for plain
    strings. Soft-fails per parameter.
    """
    try:
        param = system.get_Parameter(builtin)
    except Exception:
        return False
    if param is None or param.IsReadOnly:
        return False
    raw = "" if value is None else str(value).strip()
    if not raw:
        return False
    try:
        if param.StorageType.ToString() == "String":
            return bool(param.Set(raw))
    except Exception:
        pass
    try:
        if param.SetValueString(raw):
            return True
    except Exception:
        pass
    try:
        return bool(param.Set(float(raw)))
    except (TypeError, ValueError):
        try:
            return bool(param.Set(int(float(raw))))
        except (TypeError, ValueError):
            return False
    except Exception:
        return False


def _stamp_member_ckt_fields(group):
    panel = group.panel_name or ""
    ckt = group.circuit_token or ""
    load = group.load_name or ""
    for member in group.members:
        elem = member.element
        if elem is None:
            continue
        for name, value in (
            ("CKT_Panel_CEDT", panel),
            ("CKT_Circuit Number_CEDT", ckt),
            ("CKT_Load Name_CEDT", load),
        ):
            if not value:
                continue
            try:
                p = elem.LookupParameter(name)
            except Exception:
                p = None
            if p is None or p.IsReadOnly:
                continue
            try:
                p.Set(str(value))
            except Exception:
                continue


def execute_apply(doc, groups, panel_index=None, logger=None):
    """Apply every group sequentially. Caller manages the transaction.

    Returns a populated ``CircuitApplyResult``. Any per-group failure
    becomes a warning; the loop keeps going so partial success is the
    norm rather than the exception.
    """
    result = CircuitApplyResult()
    if not groups:
        return result
    if panel_index is None:
        panel_index = collect_panel_index(doc)
    tracker = _phasing.PanelPhaseTracker()
    for group in groups:
        if group.needs_review:
            result.skipped_count += 1
            continue
        # Phase tracker advances per group regardless of success so the
        # rotation is deterministic across re-runs.
        tracker.next_phase_for_panel(group.panel_name, group.poles)
        system, warning = apply_group(
            doc, group, panel_index,
            phase_tracker=tracker, logger=logger,
        )
        if system is None:
            result.failed_count += 1
            if warning:
                result.warnings.append("[{}] {}".format(group.key, warning))
            continue
        result.created_count += 1
        sid = _id_value(system)
        if sid is not None:
            result.system_ids_by_group_key[str(group.key)] = sid
        if warning:
            result.warnings.append("[{}] {}".format(group.key, warning))
    return result


# ---------------------------------------------------------------------
# Modeless ExternalEvent gateway
# ---------------------------------------------------------------------

class _ApplyExternalEventHandler(IExternalEventHandler):
    """Internal handler. The real work lives on ``CircuitApplyGateway``.

    ``__namespace__`` is required so pythonnet 3 registers the Python
    class with the CLR type system. Without it,
    ``ExternalEvent.Create(handler)`` raises ``"object does not
    implement IExternalEventHandler"`` because Revit's C# side checks
    ``handler is IExternalEventHandler`` and pythonnet 3's interface
    adapter doesn't satisfy that check unless the class declares its
    target namespace.

    Important: this module *must not* be cleared from ``sys.modules``
    between script runs (see ``_dev_reload.py`` — ``circuit_apply``
    is intentionally absent from the purge list). Re-importing the
    module re-executes this class statement, which registers a
    second .NET type with the same fully-qualified name and raises
    ``"Duplicate type name within an assembly"``.
    """

    __namespace__ = "MEPRFP.Automation.SuperCircuit"

    def __init__(self, gateway):
        self._gateway = gateway

    def Execute(self, uiapp):
        try:
            self._gateway._execute_pending(uiapp)
        except Exception:
            # Never raise into the Revit external event loop —
            # gateway logs the error itself.
            pass

    def GetName(self):
        return "MEPRFP SuperCircuit Apply"


class CircuitApplyGateway(object):
    """Modeless-safe wrapper around ``execute_apply``.

    Usage::

        gateway = get_or_create_gateway()
        ...
        gateway.request_apply(doc, groups, on_complete=callback)

    The window stays open while the gateway hands the work off to
    Revit's main thread via ``ExternalEvent``. ``on_complete`` is
    invoked with the populated ``CircuitApplyResult`` after the
    transaction commits (or fails).

    Use ``get_or_create_gateway()`` instead of constructing directly —
    the gateway is a per-Revit-session singleton so re-running the
    SuperCircuit pushbutton doesn't try to register a second
    ``IExternalEventHandler`` of the same fully-qualified name (which
    pythonnet 3 + the CLR refuse).
    """

    def __init__(self):
        self._handler = _ApplyExternalEventHandler(self)
        self._event = ExternalEvent.Create(self._handler)
        self._pending = None

    def request_apply(self, doc, groups, transaction_name=None,
                      on_complete=None, logger=None):
        """Queue a single apply pass. Subsequent calls before the
        previous Execute fires REPLACE the queued payload — the
        gateway is single-slot by design so a runaway click can't
        stack five identical applies.
        """
        self._pending = {
            "doc": doc,
            "groups": list(groups or []),
            "transaction_name": transaction_name or "SuperCircuit (MEPRFP 2.0)",
            "on_complete": on_complete,
            "logger": logger,
        }
        self._event.Raise()

    # ----- internal -------------------------------------------------

    def _execute_pending(self, uiapp):
        payload = self._pending
        if not payload:
            return
        self._pending = None
        from pyrevit import revit
        doc = payload["doc"]
        groups = payload["groups"]
        callback = payload["on_complete"]
        logger = payload["logger"]
        result = CircuitApplyResult()
        try:
            with revit.Transaction(payload["transaction_name"], doc=doc):
                result = execute_apply(doc, groups, logger=logger)
        except Exception as exc:
            result.warnings.append("Transaction failed: {}".format(exc))
        if callback is not None:
            try:
                callback(result)
            except Exception:
                # callback failures shouldn't crash the external event
                pass


# Module-level singleton. Surviving between SuperCircuit invocations
# means we don't need to re-call ``ExternalEvent.Create`` (which
# requires a valid Revit API context) on every run, and we sidestep
# the "Duplicate type name" pythonnet error that would arise if the
# handler class re-registered.
_GATEWAY_SINGLETON = None


def get_or_create_gateway():
    """Return the per-Revit-session ``CircuitApplyGateway``.

    First call (during a pushbutton run with a valid API context)
    constructs the gateway. Subsequent calls — including ones from
    later WPF event handlers — return the same instance so
    ``request_apply`` keeps working.

    Caller MUST be inside Revit's API execution context for the
    *first* call (i.e., during ``main()`` of a pushbutton script).
    """
    global _GATEWAY_SINGLETON
    if _GATEWAY_SINGLETON is None:
        _GATEWAY_SINGLETON = CircuitApplyGateway()
    return _GATEWAY_SINGLETON
