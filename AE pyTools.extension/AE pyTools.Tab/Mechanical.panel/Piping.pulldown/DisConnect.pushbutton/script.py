# coding: utf8
import Autodesk.Revit.DB.Plumbing as DBP
from pyrevit import script, revit, DB

from pyrevitmep.meputils import NoConnectorManagerError, get_connector_manager

logger = script.get_logger()
doc = revit.doc
output = script.get_output()

SUPPORTED_DOMAINS = {
    DB.Domain.DomainPiping: {
        "name": "Piping",
    },
    DB.Domain.DomainHvac: {
        "name": "Ducting",
    },
}


def pick_disconnect_driver(conn_a, conn_b):
    """Prefer disconnecting from MEPCurve side to preserve endpoint/system ownership."""
    try:
        a_owner = conn_a.Owner
        b_owner = conn_b.Owner
    except:
        return conn_a, conn_b

    a_is_curve = isinstance(a_owner, DB.MEPCurve)
    b_is_curve = isinstance(b_owner, DB.MEPCurve)

    if a_is_curve and not b_is_curve:
        return conn_a, conn_b
    if b_is_curve and not a_is_curve:
        return conn_b, conn_a

    # fallback: keep deterministic (disconnect from the connector we're iterating)
    return conn_a, conn_b


def disconnect_element_connectors(el, selected_ids, systems_to_divide):
    cm = get_connector_manager(el)

    for connector in cm.Connectors:
        if not connector.IsConnected:
            continue

        if connector.Domain not in SUPPORTED_DOMAINS:
            continue

        # snapshot refs (still connector objects, but we guard usage)
        try:
            refs = list(connector.AllRefs)
        except Exception as ex:
            logger.warning("Failed reading connector refs: {}".format(ex))
            continue

        for other in refs:
            if other == connector:
                continue

            try:
                other_owner = other.Owner
            except:
                continue

            # Only disconnect across selection boundary
            if other_owner and other_owner.Id.IntegerValue in selected_ids:
                continue

            # Choose disconnect direction (prefer MEPCurve side)
            driver, target = pick_disconnect_driver(connector, other)

            # Record IDs BEFORE mutation (prevents invalid object crash in print)
            try:
                driver_owner_id = driver.Owner.Id
            except:
                driver_owner_id = None

            try:
                target_owner_id = target.Owner.Id
            except:
                target_owner_id = None

            # Track the system from the driver side (carrier side)
            try:
                sys = driver.MEPSystem
                if sys:
                    systems_to_divide.add(sys)
            except:
                pass

            try:
                driver.DisconnectFrom(target)
                logger.info(
                    "Disconnected {} from {}".format(
                        driver_owner_id if driver_owner_id else "<no owner>",
                      target_owner_id if target_owner_id else "<no owner>"
                    )
                )
            except Exception as ex:
                logger.warning("Failed disconnect: {}".format(ex))



def collect_pipe_systems_from_elements(elements):
    systems = set()
    for el in elements:
        if isinstance(el, DBP.Pipe):
            try:
                for c in el.ConnectorManager.Connectors:
                    if c.MEPSystem:
                        systems.add(c.MEPSystem)
            except:
                pass
    return systems


def cleanup_orphaned_endpoints(elements):
    """
    Remove fully-disconnected endpoint elements from MEP systems
    (piping and ducting) using ConnectorSet removal.
    """
    for el in elements:
        # Skip curves — endpoints only
        if isinstance(el, DB.MEPCurve):
            continue

        try:
            cm = get_connector_manager(el)
        except:
            continue

        # Domain → {connectors, connected?, systems}
        domain_data = {}

        for c in cm.Connectors:
            domain = c.Domain
            if domain not in SUPPORTED_DOMAINS:
                continue

            data = domain_data.setdefault(domain, {
                "connectors": [],
                "connected": False,
                "systems": set(),
            })

            data["connectors"].append(c)

            if c.IsConnected:
                data["connected"] = True

            try:
                if c.MEPSystem:
                    data["systems"].add(c.MEPSystem)
            except:
                pass

        # Process each domain independently
        for domain, data in domain_data.items():
            if not data["connectors"]:
                continue

            # Still physically connected → leave it alone
            if data["connected"]:
                continue

            cset = DB.ConnectorSet()
            for c in data["connectors"]:
                try:
                    cset.Insert(c)
                except:
                    pass

            if cset.Size == 0:
                continue

            for sys in data["systems"]:
                try:
                    sys.Remove(cset)
                    logger.info(
                        "Removed orphaned {} element {} from system {}".format(
                            SUPPORTED_DOMAINS[domain]["name"],
                            el.Id,
                            sys.Id
                        )
                    )
                except Exception as ex:
                    logger.debug(
                        "Failed removing {} from system {}: {}".format(
                            el.Id, sys.Id, ex
                        )
                    )





def disconnect():
    selection = revit.get_selection()
    if not selection:
        logger.info("Nothing selected.")
        return

    selected_ids = set([el.Id.IntegerValue for el in selection])
    systems_to_divide = set()
    with DB.TransactionGroup(doc, "Disconnect Elements") as tg:
        tg.Start()
        with revit.Transaction("Disconnect elements", doc):

            # ----------------------------
            # Phase 1 — Boundary disconnect
            # ----------------------------
            for el in selection:
                try:
                    disconnect_element_connectors(el, selected_ids, systems_to_divide)
                except NoConnectorManagerError:
                    logger.warning(
                        "No connector manager found for {}: {}".format(
                            el.Category.Name, el.Id
                        )
                    )

            # ----------------------------
            # Phase 2 — Divide systems
            # ----------------------------
            for system in systems_to_divide:
                try:
                    if system.IsMultipleNetwork:
                        system.DivideSystem(doc)
                        logger.info("Divided piping system {}".format(system.Id))
                    else:
                        logger.info("System {} is still single network".format(system.Id))
                except Exception as ex:
                    logger.error("Failed dividing system: {}".format(ex))

            # ----------------------------
            # Phase 3 — Orphan cleanup
            # ----------------------------
        with revit.Transaction("Cleanup Equipment", doc):
            cleanup_orphaned_endpoints(selection)
        tg.Assimilate()


disconnect()
