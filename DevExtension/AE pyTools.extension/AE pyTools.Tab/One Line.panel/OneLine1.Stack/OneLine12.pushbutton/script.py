from pyrevit import DB, script, output
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Electrical, FamilyInstance, ElementType

# Initialize output window
out = output.get_output()
doc = __revit__.ActiveUIDocument.Document

# Initialize logger for debugging
logger = script.get_logger()


def get_parameter_value(element, param_name):
    """Retrieve the value of a parameter by name (instance and type)."""
    logger.debug("Getting parameter: %s for element ID: %s", param_name, element.Id)

    # Check for instance parameter
    param = element.LookupParameter(param_name)
    if param:
        logger.debug("Found instance parameter: %s", param_name)
        if param.StorageType == DB.StorageType.Integer:
            return param.AsInteger()

    # If not found, check for type parameter
    elem_type = doc.GetElement(element.GetTypeId())
    if elem_type:
        logger.debug("Checking type parameter for element type: %s", elem_type)
        param = elem_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if param and param.StorageType == DB.StorageType.Integer:
            return param.AsInteger()

    logger.warning("Parameter %s not found for element ID: %s", param_name, element.Id)
    return None


# Collect all FamilyInstance elements with connectors
collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment)
elements = collector.WhereElementIsNotElementType().OfClass(FamilyInstance).ToElements()

errors = []

# Loop through each element to check for connectors and systems
for elem in elements:
    logger.debug("Processing element ID: %s", elem.Id)
    mep_model = elem.MEPModel

    if mep_model and hasattr(mep_model, 'ConnectorManager') and mep_model.ConnectorManager:
        connectors = mep_model.ConnectorManager.Connectors

        # Check if connectors are present
        if connectors:
            logger.debug("Found connectors for element ID: %s", elem.Id)
            for connector in connectors:
                for ref in connector.AllRefs:
                    owner = ref.Owner
                    # Check if the owner is an ElectricalSystem
                    if isinstance(owner, Electrical.ElectricalSystem):
                        circuit = owner
                        circuit_poles = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_PHASE)

                        # Get custom parameters (instance and type)
                        poles_ced = get_parameter_value(elem, "Number of Poles_CED")
                        phase_ced = get_parameter_value(elem, "Phase_CED")

                        # Perform validation checks
                        if circuit_poles is not None and poles_ced is not None:
                            if circuit_poles != poles_ced:
                                errors.append([
                                    out.linkify(elem.Id),
                                    "Circuit poles %d != Number of Poles_CED %d" % (circuit_poles, poles_ced)
                                ])

                        if poles_ced == 3 and phase_ced != 3:
                            errors.append([
                                out.linkify(elem.Id),
                                "Expected Phase_CED to be 3, but got %d" % phase_ced
                            ])
                        elif poles_ced != 3 and phase_ced != 1:
                            errors.append([
                                out.linkify(elem.Id),
                                "Expected Phase_CED to be 1, but got %d" % phase_ced
                            ])

# Format the table data
table_data = [[error[0], error[1]] for error in errors]

# Ensure output window shows the table or a message
if table_data:
    out.print_table(
        columns=["Element ID", "Error Description"],
        table_data=table_data
    )
else:
    out.print_md("**No discrepancies found.**")

logger.info("Script execution completed")
