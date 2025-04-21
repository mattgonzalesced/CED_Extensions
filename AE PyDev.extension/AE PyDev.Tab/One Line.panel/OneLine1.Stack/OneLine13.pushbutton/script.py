# -*- coding: utf-8 -*-

__title__ = "Sync Parameters"
__doc__ = "Copies shared parameter values from equipment to corresponding symbols"

from pyrevit import revit, DB, script

# Get logger for debugging
logger = script.get_logger()

doc = revit.doc

# Mapping of shared parameters in symbols to built-in parameters in components
PARAMETER_MAP = {
    "Panel Name_CEDT": DB.BuiltInParameter.RBS_ELEC_PANEL_NAME,
    "Mains Rating_CED": DB.BuiltInParameter.RBS_ELEC_PANEL_MCB_RATING_PARAM,
    "Main Breaker Rating_CED": DB.BuiltInParameter.RBS_ELEC_PANEL_MCB_RATING_PARAM,
    "Short Circuit Rating_CEDT": DB.BuiltInParameter.RBS_ELEC_SHORT_CIRCUIT_RATING,
    "Mounting_CEDT": DB.BuiltInParameter.RBS_ELEC_MOUNTING,
    "Panel Modifications_CEDT": DB.BuiltInParameter.RBS_ELEC_MODIFICATIONS,
    "Distribution System_CEDR": DB.BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM,
    "Total Connected Load_CEDR": DB.BuiltInParameter.RBS_ELEC_PANEL_TOTALLOAD_PARAM,
    "Total Demand Load_CEDR": DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_DEMAND_CURRENT_PARAM,
    "Total Connected Current_CEDR": DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_CONNECTED_CURRENT_PARAM,
    "Total Demand Current_CEDR": DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_DEMAND_CURRENT_PARAM,
    "Voltage_CED": DB.BuiltInParameter.RBS_ELEC_VOLTAGE,
    "Number of Poles_CED": DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES,
    "CKT_Panel_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
    "CKT_Circuit Number_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER,
    "CKT_Load Name_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME,
    "CKT_Rating_CED": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM,
    "CKT_Frame_CED": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_FRAME_PARAM,
    "CKT_Wire Size_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_WIRE_SIZE_PARAM
}

# Known test elements (Component ID, Symbol ID)
test_pairs = [("3885958", "3952805"), ("3952615", "3952806")]

# Function to retrieve elements by ID
def get_element_by_id(element_id):
    """Retrieve a Revit element by its ID."""
    return doc.GetElement(DB.ElementId(int(element_id)))

# Retrieve only the test components and symbols
components = [get_element_by_id(comp_id) for comp_id, _ in test_pairs]
symbols = [get_element_by_id(sym_id) for _, sym_id in test_pairs]

# Remove None values if retrieval failed
components = [c for c in components if c]
symbols = [s for s in symbols if s]

# Log found elements
logger.info("Processing test elements only...")
for comp in components:
    logger.info("Component Found: ID {}".format(comp.Id.IntegerValue))
for sym in symbols:
    logger.info("Symbol Found: ID {}".format(sym.Id.IntegerValue))

# Create mapping from Component ID to Symbol
component_to_symbol = {}
for comp, sym in zip(components, symbols):
    component_to_symbol[comp] = sym
    logger.info("Mapping: Component {} -> Symbol {}".format(comp.Id.IntegerValue, sym.Id.IntegerValue))

# Start transaction
with revit.Transaction("Sync Parameters from Component to Symbol"):
    for component, symbol in component_to_symbol.items():
        # Get Family Symbol (Type parameters)
        family_symbol = component.Document.GetElement(component.GetTypeId())

        # Get all shared parameters for component (Instance & Type) and symbol (Instance only)
        component_params = {
            p.Definition.Name: p for p in component.Parameters if p.IsShared
        }

        if family_symbol:
            component_params.update({
                p.Definition.Name: p for p in family_symbol.Parameters if p.IsShared
            })

        symbol_params = {
            p.Definition.Name: p for p in symbol.Parameters if p.IsShared
        }

        # Find intersection of parameters (parameters that exist in both)
        common_params = set(component_params.keys()) & set(symbol_params.keys())

        # Include built-in parameters from PARAMETER_MAP
        for sym_param_name, built_in_param in PARAMETER_MAP.items():
            if sym_param_name in symbol_params:
                built_in_comp_param = component.get_Parameter(built_in_param)
                if built_in_comp_param:
                    common_params.add(sym_param_name)
                    component_params[sym_param_name] = built_in_comp_param  # Map manually

        if not common_params:
            logger.warning("No shared parameters found for Component {} and Symbol {}".format(
                component.Id.IntegerValue, symbol.Id.IntegerValue))
            continue  # Skip to next pair if no shared parameters

        # Copy values from component → symbol
        for param_name in common_params:
            comp_param = component_params.get(param_name)
            sym_param = symbol_params.get(param_name)

            if not comp_param or not sym_param:
                logger.warning("Skipping '{}' due to missing parameters.".format(param_name))
                continue

            if sym_param.StorageType == comp_param.StorageType:
                try:
                    if comp_param.StorageType == DB.StorageType.String:
                        sym_param.Set(comp_param.AsString())

                    elif comp_param.StorageType == DB.StorageType.Integer:
                        if sym_param.Definition.ParameterType == DB.ParameterType.ElementId:
                            sym_param.Set(DB.ElementId(comp_param.AsInteger()))  # Convert to ElementId if needed
                        else:
                            sym_param.Set(comp_param.AsInteger())

                    elif comp_param.StorageType == DB.StorageType.Double:
                        sym_param.Set(comp_param.AsDouble())

                    logger.info("Copied '{}' from Component {} to Symbol {}".format(
                        param_name, component.Id.IntegerValue, symbol.Id.IntegerValue))

                except Exception as e:
                    logger.warning("⚠️ Failed to copy '{}' from Component {} to Symbol {}: {}".format(
                        param_name, component.Id.IntegerValue, symbol.Id.IntegerValue, str(e)))

logger.info("✅ Parameter synchronization completed.")
