# -*- coding: utf-8 -*-
__title__ = "Get Distribution System Type"
__doc__ = """Version 1.3 - Retrieve distribution system type using AsValueString if available."""

# Import required modules
from Autodesk.Revit.DB import *
from pyrevit import script

# Get the current document
doc = __revit__.ActiveUIDocument.Document


# Helper function to get the Distribution System Type name
def get_distribution_system_type_name(panel):
    # Get the parameter containing the distribution system type (RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM)
    dist_system_param = panel.get_Parameter(BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM)

    # Check if the parameter has a value and retrieve the AsValueString
    if dist_system_param and dist_system_param.HasValue:
        dist_system_value = dist_system_param.AsValueString()

        # If AsValueString is populated, return it
        if dist_system_value:
            return dist_system_value
        else:
            # If AsValueString is not available, fall back to using ElementId (as a backup plan)
            dist_system_type_id = dist_system_param.AsElementId()
            dist_system_type = doc.GetElement(dist_system_type_id)
            if dist_system_type and hasattr(dist_system_type, 'Name'):
                return dist_system_type.Name

    return None


# Main logic
def main():
    output = script.get_output()
    output.print_md("### Distribution System Type Report")

    # Collect all electrical equipment in the project
    all_panels = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToElements()

    # Iterate through each panel and print the distribution system type name
    for panel in all_panels:
        dist_system_name = get_distribution_system_type_name(panel)

        if dist_system_name:
            output.print_md("**{}**: {}".format(panel.Name, dist_system_name))
        else:
            output.print_md("**{}**: No distribution system found".format(panel.Name))


# Run the script
if __name__ == '__main__':
    main()
