# -*- coding: utf-8 -*-


from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInParameter
from Autodesk.Revit.DB.Mechanical import MechanicalSystemType
from Autodesk.Revit.DB.Plumbing import PipingSystemType

from pyrevit import revit

# Get the active document
doc = revit.doc

# Function to get the Name parameter safely
def get_element_name_from_parameter(element):
    try:
        # Use the built-in parameter for name
        name_param = element.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM)
        if name_param:
            return name_param.AsString()
        return "Unknown Name (No Parameter)"
    except Exception as e:
        return "Unknown Name (Error: {})".format(str(e))

# Function to set the CalculationLevel parameter
def set_calculation_level(system_type, parameter_enum, level_value):
    try:
        calculation_level_param = system_type.get_Parameter(parameter_enum)
        if calculation_level_param:
            calculation_level_param.Set(level_value)
            return True
        return False
    except Exception as e:
        print("Error setting CalculationLevel: {}".format(str(e)))
        return False

# Start a transaction
transaction = Transaction(doc, "Set CalculationLevel to None")
transaction.Start()

try:
    # Collect all MechanicalSystemType elements
    mechanical_system_types = FilteredElementCollector(doc).OfClass(MechanicalSystemType)

    # Update calculation levels for MechanicalSystemType
    for system_type in mechanical_system_types:
        # Get the name of the system type
        system_name = get_element_name_from_parameter(system_type)

        # Set the CalculationLevel parameter to None (0)
        if set_calculation_level(system_type, BuiltInParameter.RBS_DUCT_SYSTEM_CALCULATION_PARAM, 0):
            print("Set Calculation Level to None for Mechanical System Type: {}".format(system_name))
        else:
            print("Failed to set Calculation Level for Mechanical System Type: {}".format(system_name))

    # Collect all PipingSystemType elements
    piping_system_types = FilteredElementCollector(doc).OfClass(PipingSystemType)

    # Update calculation levels for PipingSystemType
    for system_type in piping_system_types:
        # Get the name of the system type
        system_name = get_element_name_from_parameter(system_type)

        # Set the CalculationLevel parameter to None (0)
        if set_calculation_level(system_type, BuiltInParameter.RBS_PIPE_SYSTEM_CALCULATION_PARAM, 0):
            print("Set Calculation Level to None for Piping System Type: {}".format(system_name))
        else:
            print("Failed to set Calculation Level for Piping System Type: {}".format(system_name))

    # Commit the transaction
    transaction.Commit()

except Exception as e:
    # Roll back the transaction in case of error
    transaction.RollBack()
    print("Error: " + str(e))
