# -*- coding: utf-8 -*-

import Autodesk.Revit.DB.Plumbing as DBP
from System.Collections.Generic import List
from pyrevit import revit, DB, forms, script

from pyrevitmep.meputils import get_connector_manager, NoConnectorManagerError

doc = revit.doc
logger = script.get_logger()
output = script.get_output()
output.close_others()

PIPING_DOMAIN = DB.Domain.DomainPiping


# ============================================================
# SHARED HELPERS / PARKED EXPERIMENTS
# Some helpers are active; others are intentionally retained for reuse.
# ============================================================

def xyz_equal(a, b, tol):
    try:
        return a.DistanceTo(b) <= tol
    except:
        return False


def find_connector_by_origin(el, origin, tol):
    """
    Find a piping connector on element whose Origin matches within tolerance.
    """
    for c in iter_piping_connectors(el):
        try:
            if xyz_equal(c.Origin, origin, tol):
                return c
        except:
            continue
    return None


def _make_elementid_list(elements):
    ids = List[DB.ElementId]()
    for el in elements:
        try:
            ids.Add(el.Id)
        except:
            pass
    return ids


def _collect_connectors_from_pipes(pipes):
    """
    Return ConnectorSet for all piping connectors on given pipes.
    (Not used right now; keeping for future experiments.)
    """
    con_set = DB.ConnectorSet()
    for p in pipes:
        for c in iter_piping_connectors(p):
            try:
                con_set.Insert(c)
            except:
                pass
    return con_set


def _collect_open_connectors_from_pipes(pipes):
    """
    Collect only open piping connectors.
    (Not used in current flow.)
    """
    con_set = DB.ConnectorSet()
    for p in pipes:
        for c in iter_piping_connectors(p):
            try:
                if not c.IsConnected:
                    con_set.Insert(c)
                    return con_set
            except:
                pass
    return con_set


def _create_empty_piping_system(target_system_type):
    """
    Create an empty PipingSystem using overload available in this Revit version.
    (Not used in current flow; your tests showed it does not assign pipes reliably.)
    """
    try:
        sys_el = DBP.PipingSystem.Create(doc, target_system_type.Id)
        return sys_el
    except Exception as ex:
        logger.error("Failed creating empty piping system: {}".format(ex))
        return None


def _add_pipes_to_piping_system(piping_system, pipes):
    """
    Your version expects ConnectorSet. This previously created systems but did not
    actually assign membership in your testing.
    (Not used in current flow.)
    """
    if not piping_system or not pipes:
        return False

    connectors = _collect_open_connectors_from_pipes(pipes)
    try:
        if connectors.Size == 0:
            return False
    except:
        pass

    try:
        piping_system.Add(connectors)
        return True
    except Exception as ex:
        logger.error("Failed adding pipes to system {}: {}".format(piping_system.Id, ex))
        return False


def _create_system_for_undefined_pipes(undefined_pipes, target_system_type):
    """
    (Not used in current flow.)
    """
    if not undefined_pipes:
        return False

    print("\n--- Creating piping system for undefined pipes ---")

    new_sys = _create_empty_piping_system(target_system_type)
    if not new_sys:
        logger.error("Could not create piping system.")
        return False

    ok = _add_pipes_to_piping_system(new_sys, undefined_pipes)
    if ok:
        print("Created new system {}".format(output.linkify(new_sys.Id)))
        return True

    return False


# ============================================================
# SELECTION
# ============================================================

def get_user_selection():
    selection = revit.get_selection()
    if not selection:
        selection = revit.pick_elements(message="Select pipes / fittings / fixtures")

    elements = []
    for r in selection:
        try:
            el = doc.GetElement(r.Id)
        except:
            el = None
        if el:
            elements.append(el)

    if not elements:
        logger.error("No valid elements selected.")
        script.exit()

    return elements


def get_selected_pipes(elements):
    pipes = [el for el in elements if isinstance(el, DBP.Pipe)]
    print("Selected element count: {}".format(len(elements)))
    print("Selected pipe count: {}".format(len(pipes)))

    if not pipes:
        print("No pipes selected. Exiting (Victaulic behavior).")
        script.exit()

    return pipes


# ============================================================
# SYSTEM TYPE PICKER
# ============================================================

def pick_piping_system_type():
    system_types = (
        DB.FilteredElementCollector(doc)
        .OfClass(DBP.PipingSystemType)
        .ToElements()
    )

    type_map = {}
    for st in system_types:
        try:
            name = DB.Element.Name.__get__(st)
        except:
            name = None
        if name:
            type_map[name] = st

    selected = forms.SelectFromList.show(
        sorted(type_map.keys()),
        title="Select Piping System Type",
        multiselect=False
    )

    if not selected:
        script.exit()

    return type_map[selected]


# ============================================================
# CONNECTOR HELPERS
# ============================================================

def iter_piping_connectors(el):
    try:
        cm = get_connector_manager(el)
    except NoConnectorManagerError:
        return

    for c in cm.Connectors:
        try:
            if c.Domain == PIPING_DOMAIN:
                yield c
        except:
            continue


def pick_disconnect_driver(conn_a, conn_b):
    """
    Prefer disconnecting from MEPCurve side to preserve endpoint/system ownership.
    """
    try:
        a_owner = conn_a.Owner
        b_owner = conn_b.Owner
    except Exception as e:
        print(e)
        return conn_a, conn_b

    a_is_curve = isinstance(a_owner, DB.MEPCurve)
    b_is_curve = isinstance(b_owner, DB.MEPCurve)

    if a_is_curve and not b_is_curve:
        return conn_a, conn_b
    if b_is_curve and not a_is_curve:
        return conn_b, conn_a

    return conn_a, conn_b


# ============================================================
# NETWORK EVALUATION (UNDEFINED ELIGIBILITY CHECK)
# ============================================================

def network_is_eligible(pipes):
    """
    Based on your Systemizer observations:
      - isolated pipe or pipe-elbow-pipe (linear only) -> not eligible
      - hits a Tee (3+ refs) -> eligible
      - hits a fixture/equipment (non-pipe) -> eligible

    We walk connectivity from the selected pipes across AllRefs.
    """
    visited = set()
    queue = list(pipes)

    while queue:
        el = queue.pop()
        try:
            elid = el.Id.IntegerValue
        except:
            continue

        if elid in visited:
            continue
        visited.add(elid)

        # If we hit a non-pipe, treat as eligible (fixture/equipment/accessory/etc.)
        if not isinstance(el, DBP.Pipe):
            return True

        # Tee/cross/etc.: if any connector sees >2 refs, eligible
        # (This is a heuristic; works well with your observations.)
        for c in iter_piping_connectors(el):
            try:
                refs = list(c.AllRefs)
                if len(refs) > 2:
                    return True
            except:
                pass

        # Continue walking
        for c in iter_piping_connectors(el):
            try:
                for other in c.AllRefs:
                    try:
                        queue.append(other.Owner)
                    except:
                        pass
            except:
                pass

    return False


def _is_plumbing_fixture_owner(el):
    try:
        cat = el.Category
        if not cat:
            return False
        return cat.Id.IntegerValue == int(DB.BuiltInCategory.OST_PlumbingFixtures)
    except:
        return False


def _owner_has_three_plus_piping_connectors(el):
    count = 0
    for _ in iter_piping_connectors(el):
        count += 1
        if count >= 3:
            return True
    return False


def _connector_key(conn, tol_digits=6):
    try:
        oid = conn.Owner.Id.IntegerValue
    except:
        oid = -1

    try:
        o = conn.Origin
        return (oid, round(o.X, tol_digits), round(o.Y, tol_digits), round(o.Z, tol_digits))
    except:
        return (oid, id(conn))


def _collect_eligible_connectors_for_undefined_system(pipes):
    """
    Traverse connectivity from selected pipes and collect connectors that can
    seed a new undefined-system assignment attempt:
      - plumbing fixture connectors
      - tee/cross connectors (3+ piping connectors on owner)
    """
    visited = set()
    queued = set()
    queue = []

    for p in pipes:
        try:
            pid = p.Id.IntegerValue
        except:
            continue
        queue.append(p)
        queued.add(pid)

    con_set = DB.ConnectorSet()
    seen_keys = set()

    while queue:
        el = queue.pop(0)
        try:
            elid = el.Id.IntegerValue
        except:
            continue

        if elid in visited:
            continue
        visited.add(elid)

        is_fixture = _is_plumbing_fixture_owner(el)
        is_tee_like = _owner_has_three_plus_piping_connectors(el)

        for c in iter_piping_connectors(el):
            if is_fixture or is_tee_like:
                key = _connector_key(c)
                if key not in seen_keys:
                    try:
                        con_set.Insert(c)
                        seen_keys.add(key)
                    except:
                        pass

            try:
                refs = list(c.AllRefs)
            except:
                refs = []

            for other in refs:
                if other == c:
                    continue

                try:
                    other_owner = other.Owner
                    other_id = other_owner.Id.IntegerValue
                except:
                    continue

                if other_id not in visited and other_id not in queued:
                    queue.append(other_owner)
                    queued.add(other_id)

                if _is_plumbing_fixture_owner(other_owner) or _owner_has_three_plus_piping_connectors(other_owner):
                    key = _connector_key(other)
                    if key not in seen_keys:
                        try:
                            con_set.Insert(other)
                            seen_keys.add(key)
                        except:
                            pass

    return con_set


def _add_eligible_connectors_to_system(new_sys, connectors):
    if not new_sys:
        return False

    try:
        if connectors.Size == 0:
            logger.warning('Undefined system creation skipped: no eligible connectors found.')
            return False
    except:
        logger.warning('Undefined system creation skipped: invalid ConnectorSet.')
        return False

    try:
        new_sys.Add(connectors)
        return True
    except Exception as ex:
        logger.error('Failed adding eligible connectors to system {}: {}'.format(new_sys.Id, ex))
        return False


def _create_system_from_eligible_connectors(pipes, target_system_type):
    """
    Undefined mode assignment strategy requested by user:
    1) Create a new piping system instance
    2) Add only eligible connectors (fixtures + tee-like owners)
    """
    print('\n--- Undefined mode: create system + add eligible connectors ---')

    connectors = _collect_eligible_connectors_for_undefined_system(pipes)

    try:
        connector_count = connectors.Size
    except:
        connector_count = 0

    print('Eligible connectors found: {}'.format(connector_count))

    if connector_count == 0:
        return False

    new_sys = _create_empty_piping_system(target_system_type)
    if not new_sys:
        logger.error('Could not create new piping system for undefined network.')
        return False

    ok = _add_eligible_connectors_to_system(new_sys, connectors)
    if ok:
        print('Created undefined system {}'.format(output.linkify(new_sys.Id)))
    return ok


# ============================================================
# PHASE 0 — COLLECT BOUNDARY PAIRS + AFFECTED SYSTEM IDS
# ============================================================

def collect_boundary_work(elements):
    """
    Returns:
      selected_ids: set(int)
      saved_pairs: list(dict) with owner ids + connector origins (for reconnect)
      affected_system_ids: set(int) piping system ids to divide later
    """
    selected_ids = set([el.Id.IntegerValue for el in elements])
    saved_pairs = []
    affected_system_ids = set()

    print("\n--- Collecting boundary connections ---")

    # collect systems directly from selected elements
    for el in elements:
        try:
            sys = el.MEPSystem
            if sys:
                affected_system_ids.add(sys.Id.IntegerValue)
        except:
            pass

        for c in iter_piping_connectors(el):
            try:
                if not c.IsConnected:
                    continue
            except:
                continue

            try:
                refs = list(c.AllRefs)
            except:
                continue

            for other in refs:
                if other == c:
                    continue

                try:
                    other_owner = other.Owner
                except:
                    continue

                # disconnect only across selection boundary
                try:
                    if other_owner.Id.IntegerValue in selected_ids:
                        continue
                except:
                    continue

                driver, target = pick_disconnect_driver(c, other)

                # record affected system id BEFORE disconnect (store id only, not object)
                try:
                    sys = driver.MEPSystem
                    if sys:
                        affected_system_ids.add(sys.Id.IntegerValue)
                except:
                    pass

                # record reconnect descriptors BEFORE disconnect
                try:
                    driver_owner_id = driver.Owner.Id.IntegerValue
                    driver_origin = driver.Origin
                except:
                    continue

                try:
                    target_owner_id = target.Owner.Id.IntegerValue
                    target_origin = target.Origin
                except:
                    continue

                saved_pairs.append({
                    "a_owner_id": driver_owner_id,
                    "a_origin": driver_origin,
                    "b_owner_id": target_owner_id,
                    "b_origin": target_origin
                })

    print("Boundary connection pairs found: {}".format(len(saved_pairs)))
    print("Affected piping system ids: {}".format(len(affected_system_ids)))
    return selected_ids, saved_pairs, affected_system_ids


# ============================================================
# TX ACTIONS (small, single-purpose)
# ============================================================

def disconnect_pairs_now(saved_pairs):
    """
    Disconnect using live connector objects found by origin in the current model state.
    Note: uses origin matching strategy; do not change unless you replace the pair schema.
    """
    print("\n--- Disconnecting boundary pairs ---")
    ok = 0
    fail = 0
    tol = 0.001

    for pair in saved_pairs:
        a_el = doc.GetElement(DB.ElementId(pair["a_owner_id"]))
        b_el = doc.GetElement(DB.ElementId(pair["b_owner_id"]))
        if not a_el or not b_el:
            fail += 1
            logger.warning("Disconnect skip: missing owner element(s).")
            continue

        a_conn = find_connector_by_origin(a_el, pair["a_origin"], tol)
        b_conn = find_connector_by_origin(b_el, pair["b_origin"], tol)
        if not a_conn or not b_conn:
            fail += 1
            logger.warning("Disconnect skip: could not resolve connector(s) by origin.")
            continue

        try:
            if a_conn.IsConnected:
                a_conn.DisconnectFrom(b_conn)
            ok += 1
            print("Disconnected {} from {}".format(output.linkify(a_el.Id), output.linkify(b_el.Id)))
        except Exception as ex:
            fail += 1
            logger.warning("Failed disconnect: {}".format(ex))

    print("Disconnect results -> ok: {}, fail: {}".format(ok, fail))


def divide_affected_systems(affected_system_ids):
    print("\n--- Dividing affected systems ---")
    ok = 0
    skip = 0
    fail = 0

    created_ids = set()
    survived_ids = set()

    for sid in sorted(list(affected_system_ids)):
        try:
            sys_el = doc.GetElement(DB.ElementId(sid))
            if not sys_el:
                skip += 1
                continue

            try:
                if hasattr(sys_el, "IsValidObject") and (not sys_el.IsValidObject):
                    skip += 1
                    continue
            except:
                skip += 1
                continue

            if not isinstance(sys_el, DBP.PipingSystem):
                skip += 1
                continue

            survived_ids.add(sid)

            try:
                multi = sys_el.IsMultipleNetwork
            except:
                skip += 1
                continue

            if not multi:
                skip += 1
                print("System {} is single network".format(output.linkify(sys_el.Id)))
                continue

            try:
                new_ids = sys_el.DivideSystem(doc)
            except Exception as ex:
                fail += 1
                logger.warning("Failed dividing system {}: {}".format(sid, ex))
                continue

            ok += 1
            print("Divided system {}".format(output.linkify(DB.ElementId(sid))))

            if new_ids:
                for nid in new_ids:
                    try:
                        created_ids.add(nid.IntegerValue)
                    except:
                        pass

        except Exception as ex_outer:
            fail += 1
            logger.warning("Divide loop error on system {}: {}".format(sid, ex_outer))

    if created_ids:
        print("New systems created by DivideSystem: {}".format(len(created_ids)))

    print("Divide results -> ok: {}, skip: {}, fail: {}".format(ok, skip, fail))

    refreshed = set()
    for sid in survived_ids:
        refreshed.add(sid)
    for sid in created_ids:
        refreshed.add(sid)

    return refreshed


def _get_pipe_system(pipe):
    try:
        sys = pipe.MEPSystem
        if sys:
            return sys
    except:
        pass
    return None


def _collect_pipe_system_ids(pipes):
    sys_ids = set()
    for p in pipes:
        sys = _get_pipe_system(p)
        if sys:
            try:
                sys_ids.add(sys.Id.IntegerValue)
            except:
                pass
    return sys_ids


def _get_pipe_system_preop(pipe):
    """
    Pre-operation detection for plan routing:
    - Prefer pipe.MEPSystem
    - Fallback to connector.MEPSystem (more reliable before topology edits)
    """
    sys = _get_pipe_system(pipe)
    if sys:
        return sys

    for c in iter_piping_connectors(pipe):
        try:
            sys = c.MEPSystem
        except:
            sys = None
        if sys:
            return sys

    return None


def _collect_pipe_system_ids_preop(pipes):
    """
    System detection used for evaluation/plan selection only.
    """
    sys_ids = set()
    for p in pipes:
        sys = _get_pipe_system_preop(p)
        if sys:
            try:
                sys_ids.add(sys.Id.IntegerValue)
            except:
                pass
    return sys_ids


def _collect_pipe_system_ids_from_connectors(pipes):
    """
    Post-disconnect/divide collector.
    Connector-level MEPSystem is more dependable after topology edits.
    """
    sys_ids = set()
    for p in pipes:
        for c in iter_piping_connectors(p):
            try:
                sys = c.MEPSystem
            except:
                sys = None

            if sys:
                try:
                    sys_ids.add(sys.Id.IntegerValue)
                except:
                    pass
    return sys_ids


def _set_type_on_system(sys_el, target_system_type):
    try:
        param = sys_el.LookupParameter("Type")
        if not param:
            raise Exception("System has no 'Type' parameter.")
        param.Set(target_system_type.Id)
        return True
    except Exception as ex:
        logger.error("Failed setting type on system {}: {}".format(sys_el.Id, ex))
        return False


def _set_type_on_system_ids(system_ids, target_system_type):
    print("\n--- Setting system type on isolated systems (existing systems) ---")
    ok = 0
    fail = 0
    seen = set()

    for sid in sorted(list(system_ids)):
        if sid in seen:
            continue
        seen.add(sid)

        sys_el = doc.GetElement(DB.ElementId(sid))
        if not sys_el:
            continue
        if not isinstance(sys_el, DBP.PipingSystem):
            continue

        if _set_type_on_system(sys_el, target_system_type):
            ok += 1
            print("Set type on system {}".format(output.linkify(sys_el.Id)))
        else:
            fail += 1

    print("Set-type results -> ok: {}, fail: {}".format(ok, fail))
    return ok, fail


def _set_pipe_system_type_param(pipes, target_system_type):
    """
    For undefined networks where Revit can propagate (tee, fixture, etc.),
    this mimics Systemizer: set the pipe param and let Revit assign/propagate.
    """
    print("\n--- Setting pipe system type param (undefined propagation mode) ---")
    ok = 0
    fail = 0

    for p in pipes:
        try:
            param = p.get_Parameter(DB.BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)
            if not param:
                raise Exception("Pipe has no RBS_PIPING_SYSTEM_TYPE_PARAM.")
            param.Set(target_system_type.Id)
            ok += 1
            print("Set system type on pipe {}".format(output.linkify(p.Id)))
        except Exception as ex:
            fail += 1
            logger.error("Failed setting system type on pipe {}: {}".format(p.Id, ex))

    print("Pipe-param results -> ok: {}, fail: {}".format(ok, fail))
    return ok, fail


def reconnect_pairs(saved_pairs):
    print("\n--- Reconnecting boundary pairs ---")
    ok = 0
    fail = 0
    tol = 0.001

    for pair in saved_pairs:
        a_el = doc.GetElement(DB.ElementId(pair["a_owner_id"]))
        b_el = doc.GetElement(DB.ElementId(pair["b_owner_id"]))
        if not a_el or not b_el:
            fail += 1
            logger.warning("Reconnect skip: missing owner element(s).")
            continue

        a_conn = find_connector_by_origin(a_el, pair["a_origin"], tol)
        b_conn = find_connector_by_origin(b_el, pair["b_origin"], tol)
        if not a_conn or not b_conn:
            fail += 1
            logger.warning("Reconnect skip: could not resolve connector(s) by origin.")
            continue

        try:
            a_conn.ConnectTo(b_conn)
            ok += 1
            print("Reconnected {} to {}".format(output.linkify(a_el.Id), output.linkify(b_el.Id)))
        except Exception as ex:
            fail += 1
            logger.warning("Failed reconnect: {}".format(ex))

    print("Reconnect results -> ok: {}, fail: {}".format(ok, fail))


# ============================================================
# TRANSACTION WRAPPERS
# ============================================================

def _run_tx(tx_name, fn, *args, **kwargs):
    """
    Small wrapper so main() isn't hard-coded into 4 transactions.
    """
    with revit.Transaction(tx_name):
        return fn(*args, **kwargs)


# ============================================================
# PLAN / DECISION PHASE
# ============================================================

def _evaluate_selection(elements, pipes, saved_pairs, affected_system_ids):
    """
    Returns a dict with decision flags and computed facts.
    """
    data = {}
    data["boundary_pair_count"] = len(saved_pairs)
    data["has_boundary_pairs"] = (len(saved_pairs) > 0)

    # "Has systems" means at least one selected pipe has an actual system object
    sys_ids = _collect_pipe_system_ids_preop(pipes)
    data["selected_pipe_system_ids"] = sys_ids
    data["has_existing_systems"] = (len(sys_ids) > 0)

    # Affected systems are derived from boundary work (may be empty even if pipes have systems)
    data["affected_system_ids"] = affected_system_ids

    # Eligibility check only matters when there are no systems
    if not data["has_existing_systems"]:
        data["undefined_network_eligible"] = network_is_eligible(pipes)
    else:
        data["undefined_network_eligible"] = False

    return data


def _print_plan(eval_data):
    print("\n--- Evaluation ---")
    print("Has existing systems on selected pipes: {}".format(eval_data["has_existing_systems"]))
    print("Selected pipe system id count: {}".format(len(eval_data["selected_pipe_system_ids"])))
    print("Boundary pairs: {}".format(eval_data["boundary_pair_count"]))
    print("Affected system ids: {}".format(len(eval_data["affected_system_ids"])))
    print("Undefined network eligible: {}".format(eval_data["undefined_network_eligible"]))


# ============================================================
# MAIN (EVALUATE FIRST, THEN APPLY SEQUENCE)
# ============================================================

def main():
    print("\n=== SET PIPING SYSTEM (EVALUATE -> APPLY PLAN) ===\n")

    elements = get_user_selection()
    pipes = get_selected_pipes(elements)
    target_system_type = pick_piping_system_type()

    # Always collect boundary pairs up front (cheap + used in both plans)
    selected_ids, saved_pairs, affected_system_ids = collect_boundary_work(elements)

    eval_data = _evaluate_selection(elements, pipes, saved_pairs, affected_system_ids)
    _print_plan(eval_data)

    # --------------------------------------------------------
    # PLAN A: Existing systems present (do NOT break this flow)
    # Disconnect -> Divide -> Set system type on systems -> Reconnect
    # --------------------------------------------------------
    if eval_data["has_existing_systems"]:
        if eval_data["has_boundary_pairs"]:
            _run_tx("SetPipeSystem - Disconnect boundary", disconnect_pairs_now, saved_pairs)

        # Divide only if we actually have affected system ids
        if len(eval_data["affected_system_ids"]) > 0:
            new_affected = _run_tx("SetPipeSystem - Divide affected systems",
                                   divide_affected_systems,
                                   eval_data["affected_system_ids"])
            eval_data["affected_system_ids"] = new_affected

        # Re-collect system ids from pipes post-divide (safer than relying on affected ids)
        sys_ids_after = _collect_pipe_system_ids_from_connectors(pipes)
        _run_tx("SetPipeSystem - Set type on existing systems",
                _set_type_on_system_ids,
                sys_ids_after,
                target_system_type)

        if eval_data["has_boundary_pairs"]:
            _run_tx("SetPipeSystem - Reconnect boundary", reconnect_pairs, saved_pairs)

        print("\n=== COMPLETE (PLAN A) ===\n")
        return

    # --------------------------------------------------------
    # PLAN B: No systems on selected pipes (undefined mode)
    #
    # - If eligible (tee/fixture/etc.), create a new system and
    #   add only eligible connectors (fixture / tee-like connectors).
    #
    # - If NOT eligible (isolated linear runs), do nothing.
    # --------------------------------------------------------
    if not eval_data["undefined_network_eligible"]:
        print("\nUndefined network is not eligible for system assignment (matches Systemizer behavior).")
        print("No changes made.")
        print("\n=== COMPLETE (PLAN B - NOOP) ===\n")
        return

    created = _run_tx("SetPipeSystem - Create system from eligible connectors (undefined)",
                      _create_system_from_eligible_connectors,
                      pipes,
                      target_system_type)

    if not created:
        logger.warning("Undefined system creation did not assign any connectors.")

    print("\n=== COMPLETE (PLAN B) ===\n")


if __name__ == "__main__":
    main()
