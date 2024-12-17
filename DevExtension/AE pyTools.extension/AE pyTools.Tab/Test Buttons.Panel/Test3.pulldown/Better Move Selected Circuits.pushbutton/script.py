# -*- coding: utf-8 -*-
from pyrevit import script, forms, DB
from Autodesk.Revit.DB import (Element,
    FilteredElementCollector, Electrical, Transaction,
    BuiltInCategory, BuiltInParameter, FamilyInstance, ElementId
)

logger = script.get_logger()


class CircuitProcessor:
    """Class to handle processing of selected circuits."""
    def __init__(self, doc, include_electrical_equipment=True):
        self.doc = doc
        self.include_electrical_equipment = include_electrical_equipment
        self.circuits = []
        self.discarded_elements = []

    def process_selection(self):
        """Process the selected elements and extract circuits."""
        selection = __revit__.ActiveUIDocument.Selection.GetElementIds()
        for element_id in selection:
            element = self.doc.GetElement(element_id)
            logger.info("Processing element: %s", element_id)

            if isinstance(element, Electrical.ElectricalSystem):
                self._add_circuit(element)
            elif isinstance(element, FamilyInstance):
                self._process_family_instance(element_id, element)
            else:
                self._discard_element(element_id)

        self._finalize_results()

    def _add_circuit(self, circuit):
        """Add a circuit to the list."""
        logger.info("Found electrical system: %s", circuit.Id)
        self.circuits.append(circuit)

    def _process_family_instance(self, element_id, element):
        """Handle FamilyInstance elements."""
        if self._should_skip_equipment(element_id, element):
            return

        mep_model = element.MEPModel
        if mep_model and mep_model.ConnectorManager and mep_model.ConnectorManager.Connectors.Size > 0:
            self._process_connectors(mep_model.ConnectorManager, element_id)
        else:
            self._discard_element(element_id)

    def _should_skip_equipment(self, element_id, element):
        """Check if electrical equipment should be skipped."""
        if (
            element.Category.Id == ElementId(BuiltInCategory.OST_ElectricalEquipment)
            and not self.include_electrical_equipment
        ):
            logger.info("Skipping electrical equipment: %s", element_id)
            self.discarded_elements.append(element_id)
            return True
        return False

    def _process_connectors(self, connector_manager, element_id):
        """Process the connectors for primary circuits."""
        found_primary_circuit = False
        connector_iterator = connector_manager.Connectors.ForwardIterator()

        while connector_iterator.MoveNext():
            connector = connector_iterator.Current
            connector_info = connector.GetMEPConnectorInfo()

            if connector_info and connector_info.IsPrimary:
                found_primary_circuit = True
                logger.info("Primary connector found on family instance: %s", element_id)

                if not connector.AllRefs:
                    self._discard_element(element_id)
                else:
                    self._process_references(connector.AllRefs, element_id)
                break

        if not found_primary_circuit:
            self._discard_element(element_id)

    def _process_references(self, references, element_id):
        """Process references connected to the connector."""
        for ref in references:
            ref_owner = ref.Owner
            if ref_owner and isinstance(ref_owner, Electrical.ElectricalSystem):
                logger.info("Adding circuit from primary connector's owner: %s", ref_owner.Id)
                self.circuits.append(ref_owner)
            else:
                self._discard_element(element_id)

    def _discard_element(self, element_id):
        """Add element to the discarded list."""
        logger.info("Element not connected to circuit: %s", element_id)
        self.discarded_elements.append(element_id)

    def _finalize_results(self):
        """Finalize the results by handling alerts."""
        if not self.circuits:
            logger.info("No circuits found. Exiting script.")
            forms.alert(title="No Circuits Found", msg="No circuits found from selection.")
            script.exit()

        if self.discarded_elements:
            forms.alert(
                title="Alert",
                msg="{} incompatible elements have been discarded from selection.".format(len(self.discarded_elements))
            )


from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Electrical

class PanelProcessor:
    """Class to handle panel processing and compatibility checks."""
    def __init__(self, doc):
        self.doc = doc
        self.panels = self._collect_panels()

    def _collect_panels(self):
        """Collect all panels in the project."""
        logger.info("Collecting panels from the project.")
        return FilteredElementCollector(self.doc) \
            .OfCategory(BuiltInCategory.OST_ElectricalEquipment) \
            .WhereElementIsNotElementType() \
            .ToElements()

    def get_panel_dist_system(self, panel):
        """Returns a dictionary with the panel's distribution system name, voltage, and phase."""
        panel_data = {
            'dist_system_name': None,
            'phase': None,
            'lg_voltage': None,
            'll_voltage': None
        }

        # Ensure panel is a valid Revit element
        if not hasattr(panel, "get_Parameter"):
            raise AttributeError("The provided panel object is not a valid Revit element.")

        # Try to get the secondary distribution system
        dist_system_id = None
        secondary_dist_system_param = panel.get_Parameter(BuiltInParameter.RBS_FAMILY_CONTENT_SECONDARY_DISTRIBSYS)
        if secondary_dist_system_param and secondary_dist_system_param.HasValue:
            dist_system_id = secondary_dist_system_param.AsElementId()
            logger.info("Secondary distribution system found for panel")
        else:
            # Fallback to primary distribution system
            primary_dist_system_param = panel.get_Parameter(BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM)
            if primary_dist_system_param and primary_dist_system_param.HasValue:
                dist_system_id = primary_dist_system_param.AsElementId()
                logger.info("Primary distribution system found for panel")
            else:
                logger.warning("No distribution system found for panel")
                return panel_data

        # Retrieve the distribution system type
        dist_system_type = self.doc.GetElement(dist_system_id)
        if not dist_system_type:
            logger.warning("No distribution system type found for panel")
            return panel_data

        # Get the distribution system name
        try:
            panel_data['dist_system_name'] = Electrical.Element.Name.GetValue(dist_system_type)
        except AttributeError:
            panel_data['dist_system_name'] = "Unnamed Distribution System"
            logger.warning("Could not retrieve name for distribution system of panel")

        # Get phase and voltages
        panel_data['phase'] = getattr(dist_system_type, "ElectricalPhase", None)
        panel_data['lg_voltage'] = getattr(dist_system_type, "VoltageLineToGround", None)
        panel_data['ll_voltage'] = getattr(dist_system_type, "VoltageLineToLine", None)

        return panel_data

    def get_compatible_panels(self, circuit):
        """Find compatible panels for the given circuit."""
        logger.info("Finding compatible panels.")
        circuit_data = self._get_circuit_data(circuit)
        compatible_panels = []

        for panel in self.panels:
            panel_data = self.get_panel_dist_system(panel)
            if self._is_panel_compatible(panel_data, circuit_data):
                compatible_panels.append(panel)

        return compatible_panels

    def _get_circuit_data(self, circuit):
        """Extract data from the circuit."""
        logger.info("Getting circuit data for circuit")
        poles = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES).AsInteger()
        voltage = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE).AsDouble()
        return {'poles': poles, 'voltage': voltage}

    def _is_panel_compatible(self, panel_data, circuit_data):
        """Check if a panel is compatible with the circuit."""
        circuit_poles = circuit_data['poles']
        circuit_voltage = circuit_data['voltage']
        panel_lg_voltage = panel_data.get('lg_voltage')
        panel_ll_voltage = panel_data.get('ll_voltage')

        # Compatibility rules
        if circuit_poles == 1 and panel_lg_voltage and abs(panel_lg_voltage - circuit_voltage) < 1.0:
            return True
        if circuit_poles >= 2 and panel_ll_voltage and abs(panel_ll_voltage - circuit_voltage) < 1.0:
            return True

        return False



class CircuitMover:
    """Class to handle moving circuits to a new panel."""
    def __init__(self, doc):
        self.doc = doc

    def move_circuits(self, circuits, target_panel):
        """Move circuits to the specified target panel."""
        available_slots = self._find_open_slots(target_panel)

        if not available_slots:
            forms.alert("No available slots in the target panel.", exitscript=True)

        with Transaction(self.doc, "Move Circuits to New Panel") as trans:
            trans.Start()
            for i, circuit in enumerate(circuits):
                if i < len(available_slots):
                    slot = available_slots[i]
                    circuit.SelectPanel(target_panel)
                else:
                    forms.alert("Not enough slots in the target panel.", exitscript=True)
            self.doc.Regenerate()
            trans.Commit()

    def _find_open_slots(self, target_panel):
        """Find available slots in the target panel."""
        # Replace with actual logic to find open slots
        return list(range(1, 43))


def main():
    doc = __revit__.ActiveUIDocument.Document

    # Step 1: Process selected circuits
    circuit_processor = CircuitProcessor(doc)
    circuit_processor.process_selection()
    selected_circuits = circuit_processor.circuits

    # Step 2: Find compatible panels
    panel_processor = PanelProcessor(doc)
    compatible_panels = panel_processor.get_compatible_panels(selected_circuits[0])

    if not compatible_panels:
        forms.alert("No compatible panels found.", exitscript=True)

    panel_names = [panel.Name for panel in compatible_panels]
    target_panel_name = forms.ask_for_one_item(panel_names, title="Select Target Panel", prompt="Choose a panel:")

    if not target_panel_name:
        forms.alert("No panel selected.", exitscript=True)

    target_panel = next(panel for panel in compatible_panels if panel.Name == target_panel_name)

    # Step 3: Move circuits to the selected panel
    circuit_mover = CircuitMover(doc)
    circuit_mover.move_circuits(selected_circuits, target_panel)


if __name__ == "__main__":
    main()
