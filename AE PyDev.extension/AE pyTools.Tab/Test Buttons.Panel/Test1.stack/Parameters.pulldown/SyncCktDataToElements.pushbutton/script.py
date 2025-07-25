# -*- coding: utf-8 -*-
import System
from Autodesk.Revit.DB import Electrical
from pyrevit import revit, DB
from pyrevit import script
from pyrevit.revit import query

# Set up the output window
output = script.get_output()
output.close_others()
output.set_width(800)

logger = script.get_logger()
doc = revit.doc


class CircuitParameterSync:
    def __init__(self, parameter_mapping):
        """
        Initialize the CircuitParameterSync class.

        Args:
            parameter_mapping (dict): A dictionary mapping circuit parameters to element parameters.
        """
        self.parameter_mapping = parameter_mapping
        self.missing_params = []

    def get_param_value(self, element, param):
        """
        Get the parameter value from an element, handling BuiltInParameter, GUID, and name-based parameters.

        Args:
            element: The Revit element to query.
            param: Either a BuiltInParameter, a GUID (string), or a string representing a parameter name.
        Returns:
            The value of the parameter or None if not found.
        """
        param_obj = None

        # Check if it's a BuiltInParameter
        if isinstance(param, DB.BuiltInParameter):
            param_obj = element.get_Parameter(param)

        # Check if it's a GUID
        elif isinstance(param, str) and self.is_guid(param):
            guid = System.Guid(param)
            shared_param = DB.SharedParameterElement.Lookup(doc, guid)
            if shared_param:
                param_obj = element.get_Parameter(shared_param.Id)

        # Finally, try a name-based lookup
        elif isinstance(param, str):
            params = element.GetParameters(param)
            if params:
                param_obj = params[0]  # Take the first matching parameter

        # Return the parameter value, or None if not found
        return query.get_param_value(param_obj) if param_obj else None

    def is_guid(self, value):
        """
        Check if a given string is a valid GUID.

        Args:
            value (str): The string to check.
        Returns:
            bool: True if the string is a valid GUID, False otherwise.
        """
        try:
            # Convert string to System.Guid
            _ = System.Guid(value)
            return True
        except Exception:
            return False

    def collect_circuit_data(self):
        """
        Collect data from circuits and determine which elements need to be updated.

        Returns:
            List of dictionaries containing elements to update and their new parameter values.
        Raises:
            Exception if any parameters are not found.
        """
        circuit_collector = DB.FilteredElementCollector(doc).OfClass(DB.Electrical.ElectricalSystem).ToElements()
        elements_to_update = {}

        for circuit in circuit_collector:
            # Collect circuit parameter values
            circuit_data = {}
            for element_param, circuit_param in self.parameter_mapping.items():
                value = self.get_param_value(circuit, circuit_param)
                if value is None:
                    self.missing_params.append(circuit_param)
                circuit_data[element_param] = value

            for element in circuit.Elements:
                element_id = str(element.Id)

                # Initialize or update the dictionary for this element
                if element_id not in elements_to_update:
                    elements_to_update[element_id] = {"element": element, "update": False}

                updates_needed = elements_to_update[element_id]

                # Check if any parameters need updating
                for element_param, circuit_value in circuit_data.items():
                    element_value = self.get_param_value(element, element_param)

                    # Only update if values differ
                    if circuit_value != element_value:
                        updates_needed[element_param] = circuit_value
                        updates_needed["update"] = True

                # Remove the element if no updates are required
                if not updates_needed["update"]:
                    del elements_to_update[element_id]

        # If missing parameters, raise an exception
        if self.missing_params:
            raise Exception(
                "The following parameters could not be found and syncing was stopped:\n{}".format(
                    "\n".join(self.missing_params)
                )
            )

        # Return the dictionary values as a list
        return list(elements_to_update.values())

    def apply_updates(self, elements_to_update):
        """
        Apply updates to elements.

        Args:
            elements_to_update (list): List of dictionaries with elements and new parameter values.
        """
        with revit.Transaction("Sync Circuit Data with Connected Elements"):
            for update in elements_to_update:
                element = update["element"]
                for element_param, new_value in update.items():
                    if element_param in ["element", "update"]:
                        continue
                    param_obj = query.get_param(element, element_param)
                    if param_obj:
                        param_obj.Set(new_value)

    def output_results(self, elements_to_update):
        """
        Dynamically print results based on the number of elements to update.

        Args:
            elements_to_update: List of dictionaries containing updated elements and parameter values.
        """
        count = len(elements_to_update)

        if count > 200:
            # Simplified output for large updates
            print("More than 200 elements to update. Simplified output:")
            for update in elements_to_update:
                element_id = update["element"].Id
                updated_params = [
                    "{}: {}".format(param, update.get(param, "No Change"))
                    for param in self.parameter_mapping.keys()
                    if param in update
                ]
                print("Element ID: {}, Updates: {}".format(element_id, ", ".join(updated_params)))
        else:
            # Detailed table for smaller updates
            columns = ["Element ID"] + list(self.parameter_mapping.keys())
            table_data = []
            for update in elements_to_update:
                row = [output.linkify(update["element"].Id)]
                for param in self.parameter_mapping.keys():
                    row.append(update.get(param, "No Change"))
                table_data.append(row)

            output.print_table(
                table_data=table_data,
                title="Circuit Data Sync: Updates Summary",
                columns=columns
            )

        # Print summary
        print("Total elements updated: {}".format(count))

    def sync(self):
        """
        Main function to collect data, apply updates, and output results.
        """
        try:
            elements_to_update = self.collect_circuit_data()
            if not elements_to_update:
                print("No elements require updates. Sync process is complete.")
                return

            self.apply_updates(elements_to_update)
            print("Sync Complete: {} elements updated.".format(len(elements_to_update)))
            self.output_results(elements_to_update)

        except Exception as e:
            logger.warning(str(e))


# Define the parameter mapping
PARAMETER_MAPPING = {
    "CKT_Rating_CED": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM,
    "CKT_Load Name_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME,
    "CKT_Schedule Notes_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM,
    "CKT_Panel_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
    "CKT_Circuit Number_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER,
    "CKT_Wire Size_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_WIRE_SIZE_PARAM,

}

# Instantiate and run the sync class
if __name__ == "__main__":
    sync_tool = CircuitParameterSync(PARAMETER_MAPPING)
    sync_tool.sync()
