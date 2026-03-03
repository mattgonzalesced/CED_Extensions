# -*- coding: utf-8 -*-
"""
REFRIGERATION TAG SCRIPT - CONNECTOR BASED
===========================================
This script finds refrigeration equipment in Revit and labels them based on
their electrical connector types (primary vs secondary).

Primary connector = FAN+LTS+WMR (or FAN+LTS if no secondary exists)
Secondary connector = DEFROST

Example output:
    PANEL-A - 1, 2
    FAN+LTS+WMR
    PANEL-A - 3, 4
    DEFROST
"""

from pyrevit import revit, DB, script

# Get the current Revit document
doc = revit.doc
logger = script.get_logger()


def get_connector_type(electrical_system):
    """
    Determines if this electrical system uses primary or secondary connector.

    Returns: "primary" or "secondary"

    TODO: Update this with the actual Revit API property/method
    """
    # PLACEHOLDER - Replace with actual Revit API call
    # Options might be:
    # - electrical_system.ConnectorType
    # - electrical_system.ConnectionType
    # - electrical_system.get_Parameter(someBuiltInParameter)
    # - checking the connector's properties directly

    # For now, this is a placeholder that you'll need to update:
    # Example of what it might look like:
    """
    connector = electrical_system.BaseConnector
    if connector:
        # Maybe check connector.Domain or connector.ConnectorType
        if "primary" in str(connector.Description).lower():
            return "primary"
    """

    # TEMPORARY - YOU NEED TO REPLACE THIS
    # This is just guessing based on circuit number containing "1" or "2"
    circuit_num = electrical_system.CircuitNumber or ""
    if "1" in circuit_num or "2" in circuit_num:
        return "primary"
    return "secondary"


def get_electrical_systems_by_type(fixture):
    """
    Gets all electrical systems connected to a fixture, organized by connector type.

    Returns a dictionary like:
    {
        'PANEL-A': {
            'primary': ['1', '2'],
            'secondary': ['3', '4']
        }
    }
    """
    # Get the MEP model
    mep_model = getattr(fixture, 'MEPModel', None)
    if not mep_model:
        return {}

    # Get electrical systems
    systems = list(mep_model.GetAssignedElectricalSystems())
    if not systems:
        all_systems = mep_model.GetElectricalSystems()
        systems = [s for s in all_systems if s.GetType().Name == 'ElectricalSystem']

    if not systems:
        return {}

    # Group by panel and connector type
    panel_data = {}

    for system in systems:
        panel_name = getattr(system, 'PanelName', '') or 'NO-PANEL'
        circuit_number = system.CircuitNumber or 'NO-CIRCUIT'
        connector_type = get_connector_type(system)

        # Initialize panel data if needed
        if panel_name not in panel_data:
            panel_data[panel_name] = {
                'primary': [],
                'secondary': []
            }

        # Add circuit to appropriate connector type
        panel_data[panel_name][connector_type].append(circuit_number)

    return panel_data


def format_circuits(circuit_list):
    """
    Formats a list of circuits for display.
    Adds parentheses to circuits containing commas.
    """
    formatted = []
    for circuit in circuit_list:
        if ',' in circuit:
            formatted.append('({})'.format(circuit))
        else:
            formatted.append(circuit)
    return ', '.join(formatted)


def create_tag_text(fixture):
    """
    Creates the text that will appear in the tag based on connector types.
    """
    # Get electrical systems organized by type
    panel_data = get_electrical_systems_by_type(fixture)

    if not panel_data:
        return None

    # Build output lines
    lines = []

    for panel_name in panel_data:
        primary_circuits = panel_data[panel_name]['primary']
        secondary_circuits = panel_data[panel_name]['secondary']

        # Add primary circuits if they exist
        if primary_circuits:
            lines.append('{} - {}'.format(panel_name, format_circuits(primary_circuits)))

            # Label depends on whether secondary exists
            if secondary_circuits:
                lines.append('FAN+LTS+WMR')
            else:
                lines.append('FAN+LTS')

        # Add secondary circuits if they exist
        if secondary_circuits:
            lines.append('{} - {}'.format(panel_name, format_circuits(secondary_circuits)))
            lines.append('DEFROST')

    # Join with Windows line breaks
    return '\r\n'.join(lines)


def write_to_parameter(element, param_name, value):
    """
    Writes a value to a parameter on an element.
    """
    # Try instance parameter
    param = element.LookupParameter(param_name)
    if param and not param.IsReadOnly:
        param.Set(value)
        return True

    # Try type parameter
    try:
        element_type = element.Symbol
        param = element_type.LookupParameter(param_name)
        if param and not param.IsReadOnly:
            param.Set(value)
            return True
    except:
        pass

    return False


# ============================================================================
# MAIN SCRIPT
# ============================================================================

# Find all electrical fixtures
all_fixtures = (
    DB.FilteredElementCollector(doc)
    .OfCategory(DB.BuiltInCategory.OST_ElectricalFixtures)
    .WhereElementIsNotElementType()
)

# Filter to only "REFRIGERATION PLAN" fixtures
refrigeration_fixtures = []

for fixture in all_fixtures:
    type_id = fixture.GetTypeId()
    type_element = doc.GetElement(type_id)

    if type_element:
        name_param = type_element.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if name_param:
            type_name = name_param.AsString()
            if type_name == "REFRIGERATION PLAN":
                refrigeration_fixtures.append(fixture)

# Check if we found any
if not refrigeration_fixtures:
    logger.info("No refrigeration fixtures found.")
    script.exit()

logger.info("Found {} refrigeration fixtures".format(len(refrigeration_fixtures)))

# Update fixtures with tag text
transaction = DB.Transaction(doc, "Update Refrigeration Tags")
transaction.Start()

count = 0
for fixture in refrigeration_fixtures:
    text = create_tag_text(fixture)

    if text:
        if write_to_parameter(fixture, "Tag_Text", text):
            count += 1
        else:
            logger.warning("Could not write to fixture {}".format(fixture.Id.IntegerValue))
    else:
        logger.warning("Fixture {} has no circuits".format(fixture.Id.IntegerValue))

transaction.Commit()

logger.info("Updated {} fixtures with tag text".format(count))