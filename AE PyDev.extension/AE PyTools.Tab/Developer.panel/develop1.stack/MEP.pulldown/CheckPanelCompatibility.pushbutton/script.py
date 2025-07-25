# -*- coding: utf-8 -*-
__title__ = "Panel Compatibility Report"
__doc__ = """Version = 1.8 - Retrieve DistributionSysType for phase, voltage, and distribution system name."""

# Import required modules
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import DistributionSysType, ElectricalPhase
from pyrevit import script, DB

# Get the current document
doc = __revit__.ActiveUIDocument.Document


# Helper function to get all relevant data for the panel (distribution system type, name, phase, voltages)
def get_dist_system_data(panel):
    """Returns a dictionary containing the distribution system name, type, phase, L-G voltage, and L-L voltage."""
    dist_system_data = {
        'dist_system_name': None,
        'dist_system_type': None,
        'phase': None,
        'lg_voltage': None,
        'll_voltage': None
    }

    # Get distribution system name
    dist_system_param = panel.get_Parameter(DB.BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM)
    if dist_system_param and dist_system_param.HasValue:
        dist_system_data['dist_system_name'] = dist_system_param.AsValueString()

        # Retrieve the DistributionSysType element using the parameter
        dist_system_type_id = dist_system_param.AsElementId()
        dist_system_type = doc.GetElement(dist_system_type_id)

        if isinstance(dist_system_type, DistributionSysType):
            dist_system_data['dist_system_type'] = dist_system_type
            dist_system_data['phase'] = dist_system_type.ElectricalPhase  # ElectricalPhase property

            # Retrieve voltages
            lg_voltage = dist_system_type.VoltageLineToGround
            ll_voltage = dist_system_type.VoltageLineToLine

            # Check if the voltages exist and retrieve their actual values
            if lg_voltage:
                lg_voltage_param = lg_voltage.get_Parameter(DB.BuiltInParameter.RBS_VOLTAGETYPE_VOLTAGE_PARAM)
                dist_system_data['lg_voltage'] = lg_voltage_param.AsDouble() if lg_voltage_param else None
            if ll_voltage:
                ll_voltage_param = ll_voltage.get_Parameter(DB.BuiltInParameter.RBS_VOLTAGETYPE_VOLTAGE_PARAM)
                dist_system_data['ll_voltage'] = ll_voltage_param.AsDouble() if ll_voltage_param else None

    return dist_system_data


# Function to check panel compatibility and return the allowed pole/voltage configurations
def check_panel_compatibility(panel):
    """Check the panel's compatibility and return pole/voltage configurations."""
    dist_system_data = get_dist_system_data(panel)

    # If the distribution system type is not found, return an error message
    if not dist_system_data['dist_system_type']:
        return None, "{}: No Distribution System found.".format(panel.Name), None

    compatibility = []
    phase = dist_system_data['phase']
    lg_voltage = dist_system_data['lg_voltage']
    ll_voltage = dist_system_data['ll_voltage']

    # Determine the compatible pole/voltage configurations
    if phase == ElectricalPhase.ThreePhase:
        if lg_voltage:
            compatibility.append("1-pole (L-G Voltage: {:.1f}V)".format(lg_voltage))
        if ll_voltage:
            compatibility.append("2-pole or 3-pole (L-L Voltage: {:.1f}V)".format(ll_voltage))
    elif phase == ElectricalPhase.SinglePhase:
        if lg_voltage:
            compatibility.append("1-pole (L-G Voltage: {:.1f}V)".format(lg_voltage))
        if ll_voltage:
            compatibility.append("2-pole (L-L Voltage: {:.1f}V)".format(ll_voltage))

    if not compatibility:
        compatibility.append("No circuits can be connected to this panel.")

    return panel.Name, compatibility, dist_system_data['dist_system_name']


# Main logic for panel compatibility check
def main():
    # Collect all electrical panels in the project
    all_panels = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()

    output = script.get_output()
    output.print_md("### Panel Compatibility Report")

    # Check each panel for its compatibility
    for panel in all_panels:
        panel_name, result, dist_system_name = check_panel_compatibility(panel)
        if panel_name:
            output.print_md("#### {}".format(panel_name))
            output.print_md("Distribution System: {}".format(dist_system_name if dist_system_name else "N/A"))
            for option in result:
                output.print_md("- {}".format(option))
        else:
            output.print_md(result)


# Run the script
if __name__ == '__main__':
    main()
