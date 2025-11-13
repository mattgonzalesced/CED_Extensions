from collections import OrderedDict
from pyrevit import revit

DEFAULT_CREATE_LABEL = "SuperCircuitV3 - Create Circuits"
DEFAULT_APPLY_LABEL = "SuperCircuitV3 - Apply Circuit Data"


def run_creation(doc, groups, create_func, logger, transaction_label=None):
    created_systems = OrderedDict()
    if not groups:
        return created_systems

    label = transaction_label or DEFAULT_CREATE_LABEL
    with revit.Transaction(label):
        for group in groups:
            system = create_func(doc, group)
            if not system:
                if logger:
                    logger.warning("Circuit creation skipped for {}.".format(group.get("key")))
                continue
            created_systems[system.Id] = group

    return created_systems


def run_apply_data(doc, created_systems, apply_func, logger, transaction_label=None):
    if not created_systems:
        if logger:
            logger.info("No circuits were created.")
        return

    label = transaction_label or DEFAULT_APPLY_LABEL
    with revit.Transaction(label):
        for system_id, group in created_systems.items():
            system = doc.GetElement(system_id)
            if not system:
                if logger:
                    logger.warning("Could not locate system {} for data application.".format(system_id))
                continue
            apply_func(system, group)
