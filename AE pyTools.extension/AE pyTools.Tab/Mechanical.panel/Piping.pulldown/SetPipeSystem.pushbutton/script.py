# -*- coding: utf-8 -*-

import Autodesk.Revit.DB.Plumbing as DBP
import Autodesk.Revit.DB.Mechanical as DBM
from System.Collections.Generic import List
from pyrevit import revit, DB, forms, script

from pyrevitmep.meputils import get_connector_manager, NoConnectorManagerError

doc = revit.doc
logger = script.get_logger()
output = script.get_output()
output.close_others()

VERBOSE_LOGGING = False

try:
    if VERBOSE_LOGGING:
        if hasattr(logger, "set_verbose_mode"):
            logger.set_verbose_mode()
    else:
        if hasattr(logger, "reset_level"):
            logger.reset_level()
except Exception:
    pass

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


def find_connector_by_origin(el, origin, tol, connector_iter=None):
    """
    Find a connector on element whose Origin matches within tolerance.
    """
    iterator = connector_iter or iter_piping_connectors

    for c in iterator(el):
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

    logger.info("\n--- Creating piping system for undefined pipes ---")
    new_sys = _create_empty_piping_system(target_system_type)
    if not new_sys:
        logger.error("Could not create piping system.")
        return False

    ok = _add_pipes_to_piping_system(new_sys, undefined_pipes)
    if ok:
        _set_unique_mep_system_name(new_sys, target_system_type)
        logger.info("Created new system {}".format(output.linkify(new_sys.Id)))
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


def get_selected_network_mode(elements):
    pipes = [el for el in elements if isinstance(el, DBP.Pipe)]
    ducts = [el for el in elements if isinstance(el, DBM.Duct)]

    logger.info("Selected element count: {}".format(len(elements)))
    logger.info("Selected pipe count: {}".format(len(pipes)))
    logger.info("Selected duct count: {}".format(len(ducts)))

    if pipes and ducts:
        logger.warning("Mixed selection detected (pipes + ducts). Please select one or the other.")
        script.exit()

    if pipes:
        piping_elements = []
        for el in elements:
            has_piping = False
            for _ in iter_piping_connectors(el):
                has_piping = True
                break
            if has_piping:
                piping_elements.append(el)
        return "piping", piping_elements, pipes

    if ducts:
        duct_elements = []
        for el in elements:
            has_duct = False
            for _ in iter_duct_connectors(el):
                has_duct = True
                break
            if has_duct:
                duct_elements.append(el)
        return "duct", duct_elements, ducts

    logger.info("No pipes or ducts selected. Exiting.")
    script.exit()

    if pipes:
        return "piping", pipes

    if ducts:
        return "duct", ducts

    logger.info("No pipes or ducts selected. Exiting.")
    script.exit()


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


def pick_duct_system_type():
    system_types = (
        DB.FilteredElementCollector(doc)
        .OfClass(DBM.MechanicalSystemType)
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
        title="Select Duct System Type",
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


def iter_duct_connectors(el):
    try:
        cm = get_connector_manager(el)
    except NoConnectorManagerError:
        return

    for c in cm.Connectors:
        try:
            if c.Domain == DB.Domain.DomainHvac:
                yield c
        except:
            continue


def _get_duct_system(duct):
    try:
        sys = duct.MEPSystem
        if sys:
            return sys
    except:
        pass
    return None


def _get_duct_system_preop(duct):
    sys = _get_duct_system(duct)
    if sys:
        return sys

    for c in iter_duct_connectors(duct):
        try:
            sys = c.MEPSystem
        except:
            sys = None
        if sys:
            return sys

    return None


def _collect_duct_system_ids_preop(ducts):
    sys_ids = set()
    for d in ducts:
        sys = _get_duct_system_preop(d)
        if sys:
            try:
                sys_ids.add(sys.Id.IntegerValue)
            except:
                pass
    return sys_ids


def _collect_duct_system_ids_from_connectors(ducts):
    sys_ids = set()
    for d in ducts:
        for c in iter_duct_connectors(d):
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


def _set_duct_system_type_param(ducts, target_system_type):
    logger.info("\n--- Setting duct system type param ---")
    ok = 0
    fail = 0

    for d in ducts:
        try:
            param = d.get_Parameter(DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)
            if not param:
                raise Exception("Duct has no RBS_DUCT_SYSTEM_TYPE_PARAM.")
            param.Set(target_system_type.Id)
            ok += 1
            logger.info("Set system type on duct {}".format(output.linkify(d.Id)))
        except Exception as ex:
            fail += 1
            logger.error("Failed setting system type on duct {}: {}".format(d.Id, ex))

    logger.info("Duct-param results -> ok: {}, fail: {}".format(ok, fail))
    return ok, fail


def pick_disconnect_driver(conn_a, conn_b):
    """
    Prefer disconnecting from MEPCurve side to preserve endpoint/system ownership.
    """
    try:
        a_owner = conn_a.Owner
        b_owner = conn_b.Owner
    except Exception as e:
        logger.info(e)
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

    Tee handling note:
    We only capture the *branch-facing* tee connector encountered from the walk,
    not every connector on the tee owner.
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
    fixture_owner_added = set()

    while queue:
        el = queue.pop(0)
        try:
            elid = el.Id.IntegerValue
        except:
            continue

        if elid in visited:
            continue
        visited.add(elid)

        for c in iter_piping_connectors(el):
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

                # If neighboring owner is eligible, add boundary connector only.
                # Fixture rule: add only the first connected fixture connector we hit
                # per fixture owner to avoid grabbing unused fixture ports.
                if _is_plumbing_fixture_owner(other_owner):
                    if other_id in fixture_owner_added:
                        continue

                    key = _connector_key(other)
                    if key not in seen_keys:
                        try:
                            con_set.Insert(other)
                            seen_keys.add(key)
                            fixture_owner_added.add(other_id)
                        except:
                            pass
                    continue

                if _owner_has_three_plus_piping_connectors(other_owner):
                    key = _connector_key(other)
                    if key not in seen_keys:
                        try:
                            con_set.Insert(other)
                            seen_keys.add(key)
                        except:
                            pass

    return con_set


def _add_eligible_connectors_to_system(new_sys, connectors):
    """
    Add connectors one-by-one so bad connectors (already used / incompatible
    classification) do not fail the entire undefined assignment attempt.
    Returns tuple: (added_count, skipped_count)
    """
    if not new_sys:
        return 0, 0

    try:
        if connectors.Size == 0:
            logger.warning('Undefined system creation skipped: no eligible connectors found.')
            return 0, 0
    except:
        logger.warning('Undefined system creation skipped: invalid ConnectorSet.')
        return 0, 0

    added = 0
    skipped = 0
    first_skip = None

    for conn in connectors:
        one = DB.ConnectorSet()
        try:
            one.Insert(conn)
        except:
            skipped += 1
            continue

        try:
            new_sys.Add(one)
            added += 1
        except Exception as ex:
            skipped += 1
            if first_skip is None:
                first_skip = ex

    if skipped > 0:
        logger.warning('Skipped {} connector(s) during undefined add on {}. First error: {}'.format(
            skipped,
            output.linkify(new_sys.Id),
            first_skip
        ))

    return added, skipped


def _iter_piping_system_types_for_seed(target_system_type):
    """
    Yield seed system types with target first, then all other types.
    This allows fallback seeding when target classification is too strict for
    the first connector assignment on undefined networks.
    """
    yielded = set()

    if target_system_type:
        try:
            yielded.add(target_system_type.Id.IntegerValue)
            yield target_system_type
        except:
            pass

    for st in DB.FilteredElementCollector(doc).OfClass(DBP.PipingSystemType).ToElements():
        try:
            sid = st.Id.IntegerValue
        except:
            continue
        if sid in yielded:
            continue
        yielded.add(sid)
        yield st


def _seed_undefined_system_with_fallback(connectors, target_system_type):
    """
    Try target type first; if connector-domain/classification blocks assignment,
    try other seed types until at least one connector is accepted.
    """
    for seed_type in _iter_piping_system_types_for_seed(target_system_type):
        seed_name = None
        try:
            seed_name = DB.Element.Name.__get__(seed_type)
        except:
            seed_name = str(seed_type.Id.IntegerValue)

        new_sys = _create_empty_piping_system(seed_type)
        if not new_sys:
            continue

        added, skipped = _add_eligible_connectors_to_system(new_sys, connectors)
        logger.info('Undefined seed attempt [{}] -> added: {}, skipped: {}'.format(seed_name, added, skipped))
        if added > 0:
            # If we seeded with a non-target type, flip to requested target type now.
            try:
                if seed_type.Id.IntegerValue != target_system_type.Id.IntegerValue:
                    _set_type_on_system(new_sys, target_system_type)
            except:
                pass
            return True, new_sys

        # No connectors added for this seed type; delete empty system and try next.
        try:
            doc.Delete(new_sys.Id)
        except:
            pass

    return False, None


def _create_system_from_eligible_connectors(pipes, target_system_type):
    """
    Undefined mode assignment strategy requested by user:
    1) Create a new piping system instance
    2) Add only eligible connectors (fixtures + tee-like owners)
    3) Use seed-type fallback to avoid connector classification dead-ends,
       then set final type to user target.
    """
    logger.info("\n--- Undefined mode: create system + add eligible connectors ---")
    connectors = _collect_eligible_connectors_for_undefined_system(pipes)

    try:
        connector_count = connectors.Size
    except:
        connector_count = 0

    logger.info('Eligible connectors found: {}'.format(connector_count))
    if connector_count == 0:
        return False

    ok, new_sys = _seed_undefined_system_with_fallback(connectors, target_system_type)
    if ok and new_sys:
        _set_unique_mep_system_name(new_sys, target_system_type)
        logger.info('Created undefined system {}'.format(output.linkify(new_sys.Id)))
    return ok


# ============================================================
# PHASE 0 — COLLECT BOUNDARY PAIRS + AFFECTED SYSTEM IDS
# ============================================================

def collect_boundary_work(elements, connector_iter):
    """
    Returns:
      selected_ids: set(int)
      saved_pairs: list(dict) with owner ids + connector origins (for reconnect)
      affected_system_ids: set(int) system ids to divide later
    """
    selected_ids = set([el.Id.IntegerValue for el in elements])
    saved_pairs = []
    affected_system_ids = set()

    logger.info("\n--- Collecting boundary connections ---")

    for el in elements:
        try:
            sys = el.MEPSystem
            if sys:
                affected_system_ids.add(sys.Id.IntegerValue)
        except:
            pass

        for c in connector_iter(el):
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

                try:
                    if other_owner.Id.IntegerValue in selected_ids:
                        continue
                except:
                    continue

                driver, target = pick_disconnect_driver(c, other)

                try:
                    sys = driver.MEPSystem
                    if sys:
                        affected_system_ids.add(sys.Id.IntegerValue)
                except:
                    pass

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

    logger.info("Boundary connection pairs found: {}".format(len(saved_pairs)))
    logger.info("Affected system ids: {}".format(len(affected_system_ids)))
    return selected_ids, saved_pairs, affected_system_ids

def disconnect_pairs_now(saved_pairs, connector_iter):
    """
    Disconnect using live connector objects found by origin in current model state.
    """
    logger.info("\n--- Disconnecting boundary pairs ---")
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

        a_conn = find_connector_by_origin(a_el, pair["a_origin"], tol, connector_iter)
        b_conn = find_connector_by_origin(b_el, pair["b_origin"], tol, connector_iter)
        if not a_conn or not b_conn:
            fail += 1
            logger.warning("Disconnect skip: could not resolve connector(s) by origin.")
            continue

        try:
            if a_conn.IsConnected:
                a_conn.DisconnectFrom(b_conn)
            ok += 1
            logger.info("Disconnected {} from {}".format(output.linkify(a_el.Id), output.linkify(b_el.Id)))
        except Exception as ex:
            fail += 1
            logger.warning("Failed disconnect: {}".format(ex))

    logger.info("Disconnect results -> ok: {}, fail: {}".format(ok, fail))

def divide_affected_systems(affected_system_ids, system_class):
    logger.info("\n--- Dividing affected systems ---")
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

            if not isinstance(sys_el, system_class):
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
                logger.info("System {} is single network".format(output.linkify(sys_el.Id)))
                continue

            try:
                new_ids = sys_el.DivideSystem(doc)
            except Exception as ex:
                fail += 1
                logger.warning("Failed dividing system {}: {}".format(sid, ex))
                continue

            ok += 1
            logger.info("Divided system {}".format(output.linkify(DB.ElementId(sid))))

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
        logger.info("New systems created by DivideSystem: {}".format(len(created_ids)))
    logger.info("Divide results -> ok: {}, skip: {}, fail: {}".format(ok, skip, fail))

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


def _get_pipe_level_id(pipe):
    try:
        lvl = pipe.ReferenceLevel
        if lvl:
            return lvl.Id
    except:
        pass

    for bip in [DB.BuiltInParameter.RBS_START_LEVEL_PARAM,
                DB.BuiltInParameter.LEVEL_PARAM]:
        try:
            p = pipe.get_Parameter(bip)
            if p:
                lid = p.AsElementId()
                if lid and lid != DB.ElementId.InvalidElementId:
                    return lid
        except:
            pass

    return None


def _get_single_open_pipe_connector(pipe):
    open_connectors = []
    for c in iter_piping_connectors(pipe):
        try:
            if not c.IsConnected:
                open_connectors.append(c)
        except:
            pass

    if len(open_connectors) == 1:
        return open_connectors[0]

    return None


def _create_short_pipe_from_open_connector(pipe, open_conn, target_system_type):
    length_ft = 0.5 / 12.0

    try:
        start = open_conn.Origin
    except Exception as ex:
        logger.error("Failed to read open connector origin: {}".format(ex))
        return False

    direction = None
    try:
        curve = pipe.Location.Curve
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        if p0.DistanceTo(start) <= p1.DistanceTo(start):
            direction = (p0 - p1).Normalize()
        else:
            direction = (p1 - p0).Normalize()
    except:
        pass

    if not direction:
        try:
            direction = open_conn.CoordinateSystem.BasisZ.Normalize()
        except Exception as ex:
            logger.error("Failed to resolve direction for stub pipe: {}".format(ex))
            return False

    end = start + (direction.Multiply(length_ft))

    pipe_type_id = pipe.GetTypeId()
    level_id = _get_pipe_level_id(pipe)
    if not level_id:
        logger.error("Could not determine level for selected pipe {}".format(output.linkify(pipe.Id)))
        return False

    try:
        stub = DBP.Pipe.Create(doc, target_system_type.Id, pipe_type_id, level_id, start, end)
    except Exception as ex:
        logger.error("Failed creating short pipe stub: {}".format(ex))
        return False

    try:
        src_dia = pipe.get_Parameter(DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
        dst_dia = stub.get_Parameter(DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
        if src_dia and dst_dia:
            dst_dia.Set(src_dia.AsDouble())
    except Exception as ex:
        logger.warning("Could not copy diameter to short pipe stub: {}".format(ex))

    try:
        nearest = None
        nearest_d = None
        for c in iter_piping_connectors(stub):
            d = c.Origin.DistanceTo(start)
            if nearest is None or d < nearest_d:
                nearest = c
                nearest_d = d
        if nearest:
            open_conn.ConnectTo(nearest)
    except Exception as ex:
        logger.warning("Could not connect short pipe stub to source connector: {}".format(ex))

    try:
        new_sys = stub.MEPSystem
        if new_sys:
            _set_unique_mep_system_name(new_sys, target_system_type)
    except:
        pass

    logger.info("Created short pipe stub {} from {}".format(output.linkify(stub.Id), output.linkify(pipe.Id)))
    return True


def _is_selected_pipe_element(el, selected_pipe_ids):
    try:
        return isinstance(el, DBP.Pipe) and el.Id.IntegerValue in selected_pipe_ids
    except:
        return False


def _open_connector_has_selected_path_to_tee(open_conn, selected_pipe_ids):
    try:
        start_owner = open_conn.Owner
        start_id = start_owner.Id.IntegerValue
    except:
        return False

    if start_id not in selected_pipe_ids:
        return False

    queue = [start_owner]
    visited = set()

    while queue:
        el = queue.pop(0)
        try:
            eid = el.Id.IntegerValue
        except:
            continue

        if eid in visited:
            continue
        visited.add(eid)

        for c in iter_piping_connectors(el):
            try:
                refs = list(c.AllRefs)
            except:
                refs = []

            for other in refs:
                if other == c:
                    continue

                try:
                    other_owner = other.Owner
                except:
                    continue

                if _owner_has_three_plus_piping_connectors(other_owner):
                    return True

                if _is_selected_pipe_element(other_owner, selected_pipe_ids):
                    try:
                        oid = other_owner.Id.IntegerValue
                    except:
                        continue
                    if oid not in visited:
                        queue.append(other_owner)

    return False


def _create_open_branch_stubs_from_selection(pipes, target_system_type):
    selected_pipe_ids = set()
    for p in pipes:
        try:
            selected_pipe_ids.add(p.Id.IntegerValue)
        except:
            pass

    if not selected_pipe_ids:
        return 0

    created = 0
    seen = set()

    for pipe in pipes:
        for conn in iter_piping_connectors(pipe):
            try:
                if conn.IsConnected:
                    continue
            except:
                continue

            key = _connector_key(conn)
            if key in seen:
                continue
            seen.add(key)

            if not _open_connector_has_selected_path_to_tee(conn, selected_pipe_ids):
                continue

            if _create_short_pipe_from_open_connector(pipe, conn, target_system_type):
                created += 1

    return created


def _get_system_type_abbreviation(system_type):
    try:
        p = system_type.LookupParameter("Abbreviation")
        if p:
            val = p.AsString()
            if val:
                return val.strip()
    except:
        pass

    try:
        p = system_type.get_Parameter(DB.BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM)
        if p:
            val = p.AsString()
            if val:
                return val.strip()
    except:
        pass

    try:
        n = DB.Element.Name.__get__(system_type)
        if n:
            return n.split()[0]
    except:
        pass

    return "SYS"


def _collect_used_mep_system_names():
    names = set()

    for s in (DB.FilteredElementCollector(doc)
              .OfClass(DBP.PipingSystem)
              .WhereElementIsNotElementType()
              .ToElements()):
        try:
            n = DB.Element.Name.__get__(s)
            if n:
                names.add(n)
        except:
            pass

    for s in (DB.FilteredElementCollector(doc)
              .OfClass(DBM.MechanicalSystem)
              .WhereElementIsNotElementType()
              .ToElements()):
        try:
            n = DB.Element.Name.__get__(s)
            if n:
                names.add(n)
        except:
            pass

    return names


def _set_unique_mep_system_name(system_el, target_system_type):
    if not system_el:
        return False

    used = _collect_used_mep_system_names()

    try:
        current = DB.Element.Name.__get__(system_el)
        if current in used:
            used.remove(current)
    except:
        pass

    prefix = _get_system_type_abbreviation(target_system_type)

    i = 1
    while i < 100000:
        candidate = "{} {}".format(prefix, i)
        if candidate not in used:
            try:
                system_el.Name = candidate
                return True
            except:
                try:
                    DB.Element.Name.__set__(system_el, candidate)
                    return True
                except Exception as ex:
                    logger.warning("Could not set system name {}: {}".format(candidate, ex))
                    return False
        i += 1

    logger.warning("Could not find unique system name for prefix {}".format(prefix))
    return False


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


def _set_type_on_system_ids(system_ids, target_system_type, system_class, label):
    logger.info("\n--- Setting system type on isolated {} systems ---".format(label))
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
        if not isinstance(sys_el, system_class):
            continue

        if _set_type_on_system(sys_el, target_system_type):
            _set_unique_mep_system_name(sys_el, target_system_type)
            ok += 1
            logger.info("Set type on system {}".format(output.linkify(sys_el.Id)))
        else:
            fail += 1

    logger.info("Set-type results -> ok: {}, fail: {}".format(ok, fail))
    return ok, fail


def _set_pipe_system_type_param(pipes, target_system_type):
    """
    For undefined networks where Revit can propagate (tee, fixture, etc.),
    this mimics Systemizer: set the pipe param and let Revit assign/propagate.
    """
    logger.info("\n--- Setting pipe system type param (undefined propagation mode) ---")
    ok = 0
    fail = 0

    for p in pipes:
        try:
            param = p.get_Parameter(DB.BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)
            if not param:
                raise Exception("Pipe has no RBS_PIPING_SYSTEM_TYPE_PARAM.")
            param.Set(target_system_type.Id)
            ok += 1
            logger.info("Set system type on pipe {}".format(output.linkify(p.Id)))
        except Exception as ex:
            fail += 1
            logger.error("Failed setting system type on pipe {}: {}".format(p.Id, ex))

    logger.info("Pipe-param results -> ok: {}, fail: {}".format(ok, fail))
    return ok, fail


def reconnect_pairs(saved_pairs, connector_iter):
    logger.info("\n--- Reconnecting boundary pairs ---")
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

        a_conn = find_connector_by_origin(a_el, pair["a_origin"], tol, connector_iter)
        b_conn = find_connector_by_origin(b_el, pair["b_origin"], tol, connector_iter)
        if not a_conn or not b_conn:
            fail += 1
            logger.warning("Reconnect skip: could not resolve connector(s) by origin.")
            continue

        try:
            a_conn.ConnectTo(b_conn)
            ok += 1
            logger.info("Reconnected {} to {}".format(output.linkify(a_el.Id), output.linkify(b_el.Id)))
        except Exception as ex:
            fail += 1
            logger.warning("Failed reconnect: {}".format(ex))

    logger.info("Reconnect results -> ok: {}, fail: {}".format(ok, fail))

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
    logger.info("\n--- Evaluation ---")
    logger.info("Has existing systems on selected pipes: {}".format(eval_data["has_existing_systems"]))
    logger.info("Selected pipe system id count: {}".format(len(eval_data["selected_pipe_system_ids"])))
    logger.info("Boundary pairs: {}".format(eval_data["boundary_pair_count"]))
    logger.info("Affected system ids: {}".format(len(eval_data["affected_system_ids"])))
    logger.info("Undefined network eligible: {}".format(eval_data["undefined_network_eligible"]))
# ============================================================
# MAIN (EVALUATE FIRST, THEN APPLY SEQUENCE)
# ============================================================

def main():
    logger.info("\n=== SET MEP SYSTEM (EVALUATE -> APPLY PLAN) ===\n")
    elements = get_user_selection()
    mode, mode_elements, selected_curves = get_selected_network_mode(elements)

    if mode == "piping":
        pipes = selected_curves
        target_system_type = pick_piping_system_type()

        selected_ids, saved_pairs, affected_system_ids = collect_boundary_work(mode_elements, iter_piping_connectors)
        eval_data = _evaluate_selection(mode_elements, pipes, saved_pairs, affected_system_ids)
        _print_plan(eval_data)

        with DB.TransactionGroup(doc, "SetPipeSystem - Apply Plan") as tg:
            tg.Start()

            if eval_data["has_existing_systems"]:
                if eval_data["has_boundary_pairs"]:
                    _run_tx("SetPipeSystem - Disconnect boundary", disconnect_pairs_now, saved_pairs, iter_piping_connectors)

                if len(eval_data["affected_system_ids"]) > 0:
                    new_affected = _run_tx("SetPipeSystem - Divide affected systems",
                                           divide_affected_systems,
                                           eval_data["affected_system_ids"],
                                           DBP.PipingSystem)
                    eval_data["affected_system_ids"] = new_affected

                sys_ids_after = _collect_pipe_system_ids_from_connectors(pipes)
                _run_tx("SetPipeSystem - Set type on existing systems",
                        _set_type_on_system_ids,
                        sys_ids_after,
                        target_system_type,
                        DBP.PipingSystem,
                        "piping")

                if eval_data["has_boundary_pairs"]:
                    _run_tx("SetPipeSystem - Reconnect boundary", reconnect_pairs, saved_pairs, iter_piping_connectors)

                tg.Assimilate()
                logger.info("\n=== COMPLETE (PIPING PLAN A) ===\n")
                return

            created_stubs = _run_tx("SetPipeSystem - Create open-branch stubs",
                                    _create_open_branch_stubs_from_selection,
                                    pipes,
                                    target_system_type)
            if created_stubs > 0:
                logger.info("Created {} open-branch stub pipe(s).".format(created_stubs))

            if not eval_data["undefined_network_eligible"]:
                logger.info("\nUndefined network is not eligible for system assignment (matches Systemizer behavior).")
                logger.info("No changes made.")
                tg.Assimilate()
                logger.info("\n=== COMPLETE (PIPING PLAN B - NOOP) ===\n")
                return

            created = _run_tx("SetPipeSystem - Create system from eligible connectors (undefined)",
                              _create_system_from_eligible_connectors,
                              pipes,
                              target_system_type)

            if not created:
                logger.warning("Undefined system creation did not assign any connectors.")

            tg.Assimilate()
            logger.info("\n=== COMPLETE (PIPING PLAN B) ===\n")
            return

    ducts = selected_curves
    target_system_type = pick_duct_system_type()

    selected_ids, saved_pairs, affected_system_ids = collect_boundary_work(mode_elements, iter_duct_connectors)
    sys_ids_pre = _collect_duct_system_ids_preop(ducts)

    with DB.TransactionGroup(doc, "SetDuctSystem - Apply Plan") as tg:
        tg.Start()

        if len(sys_ids_pre) > 0:
            if len(saved_pairs) > 0:
                _run_tx("SetDuctSystem - Disconnect boundary", disconnect_pairs_now, saved_pairs, iter_duct_connectors)

            if len(affected_system_ids) > 0:
                _run_tx("SetDuctSystem - Divide affected systems",
                        divide_affected_systems,
                        affected_system_ids,
                        DBM.MechanicalSystem)

            sys_ids_after = _collect_duct_system_ids_from_connectors(ducts)
            _run_tx("SetDuctSystem - Set type on existing systems",
                    _set_type_on_system_ids,
                    sys_ids_after,
                    target_system_type,
                    DBM.MechanicalSystem,
                    "duct")

            if len(saved_pairs) > 0:
                _run_tx("SetDuctSystem - Reconnect boundary", reconnect_pairs, saved_pairs, iter_duct_connectors)

            tg.Assimilate()
            logger.info("\n=== COMPLETE (DUCT EXISTING) ===\n")
            return

        _run_tx("SetDuctSystem - Set type via duct parameter",
                _set_duct_system_type_param,
                ducts,
                target_system_type)

        sys_ids_after = _collect_duct_system_ids_from_connectors(ducts)
        if len(sys_ids_after) > 0:
            _run_tx("SetDuctSystem - Name resulting systems",
                    _set_type_on_system_ids,
                    sys_ids_after,
                    target_system_type,
                    DBM.MechanicalSystem,
                    "duct")

        tg.Assimilate()
        logger.info("\n=== COMPLETE (DUCT PARAM) ===\n")

if __name__ == "__main__":
    main()
