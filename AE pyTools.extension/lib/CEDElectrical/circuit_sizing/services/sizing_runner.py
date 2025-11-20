from pyrevit import DB

from CEDElectrical.circuit_sizing.domain.circuit_branch import CircuitBranch
from CEDElectrical.circuit_sizing.services.override_validator import OverrideValidator


class CircuitSizingRunner(object):
    """Coordinates circuit sizing calculations and writes results back to Revit."""

    def __init__(self, doc, logger):
        self.doc = doc
        self.logger = logger
        self.validator = OverrideValidator(logger)

    def collect_shared_param_values(self, branch):
        return {
            'CKT_Circuit Type_CEDT': branch.branch_type,
            'CKT_Panel_CEDT': branch.panel,
            'CKT_Circuit Number_CEDT': branch.circuit_number,
            'CKT_Load Name_CEDT': branch.load_name,
            'CKT_Rating_CED': branch.rating,
            'CKT_Frame_CED': branch.frame,
            'CKT_Length_CED': branch.length,
            'CKT_Schedule Notes_CEDT': branch.circuit_notes,
            'Voltage Drop Percentage_CED': branch.voltage_drop_percentage,
            'CKT_Wire Hot Size_CEDT': branch.hot_wire_size,
            'CKT_Number of Wires_CED': branch.number_of_wires,
            'CKT_Number of Sets_CED': branch.number_of_sets,
            'CKT_Wire Hot Quantity_CED': branch.hot_wire_quantity,
            'CKT_Wire Ground Size_CEDT': branch.ground_wire_size,
            'CKT_Wire Ground Quantity_CED': branch.ground_wire_quantity,
            'CKT_Wire Neutral Size_CEDT': branch.neutral_wire_size,
            'CKT_Wire Neutral Quantity_CED': branch.neutral_wire_quantity,
            'CKT_Wire Isolated Ground Size_CEDT': branch.isolated_ground_wire_size,
            'CKT_Wire Isolated Ground Quantity_CED': branch.isolated_ground_wire_quantity,
            'Wire Material_CEDT': branch.wire_material,
            'Wire Temparature Rating_CEDT': branch.wire_info.get('wire_temperature_rating', '75'),
            'Wire Insulation_CEDT': branch.wire_info.get('wire_insulation', 'THWN'),
            'Conduit Size_CEDT': branch.conduit_size,
            'Conduit Type_CEDT': branch.conduit_type,
            'Conduit Fill Percentage_CED': branch.conduit_fill_percentage,
            'Wire Size_CEDT': branch.get_wire_size_callout(),
            'Conduit and Wire Size_CEDT': branch.get_conduit_and_wire_size(),
            'Circuit Load Current_CED': branch.circuit_load_current,
            'Circuit Ampacity_CED': branch.circuit_base_ampacity,
        }

    def update_circuit_parameters(self, circuit, param_values):
        for param_name, value in param_values.items():
            if value is None:
                continue
            param = circuit.LookupParameter(param_name)
            if not param:
                continue
            try:
                if param.StorageType == DB.StorageType.String:
                    param.Set(str(value))
                elif param.StorageType == DB.StorageType.Integer:
                    param.Set(int(value))
                elif param.StorageType == DB.StorageType.Double:
                    param.Set(float(value))
            except Exception as e:
                if self.logger:
                    self.logger.debug("❌ Failed to write '{}' to circuit {}: {}".format(param_name, circuit.Id, e))

    def update_connected_elements(self, branch, param_values):
        circuit = branch.circuit
        fixture_count = 0
        equipment_count = 0

        for el in circuit.Elements:
            if not isinstance(el, DB.FamilyInstance):
                continue

            cat = el.Category
            if not cat:
                continue

            cat_id = cat.Id
            is_fixture = cat_id == DB.ElementId(DB.BuiltInCategory.OST_ElectricalFixtures)
            is_equipment = cat_id == DB.ElementId(DB.BuiltInCategory.OST_ElectricalEquipment)

            if not (is_fixture or is_equipment):
                continue

            for param_name, value in param_values.items():
                if value is None:
                    continue
                param = el.LookupParameter(param_name)
                if not param:
                    continue
                try:
                    if param.StorageType == DB.StorageType.String:
                        param.Set(str(value))
                    elif param.StorageType == DB.StorageType.Integer:
                        param.Set(int(value))
                    elif param.StorageType == DB.StorageType.Double:
                        param.Set(float(value))
                except Exception as e:
                    if self.logger:
                        self.logger.debug("❌ Failed to write '{}' to element {}: {}".format(param_name, el.Id, e))

            if is_fixture:
                fixture_count += 1
            elif is_equipment:
                equipment_count += 1

        return fixture_count, equipment_count

    def calculate_and_update(self, circuits):
        count = len(circuits)
        if count > 1000:
            from pyrevit import forms

            proceed = forms.alert(
                "{} circuits selected.\n\nThis may take a while.\n\n".format(count),
                title="⚠️ Large Selection Warning",
                options=["Continue", "Cancel"]
            )
            if proceed != "Continue":
                return

        branches = []
        total_fixtures = 0
        total_equipment = 0

        for circuit in circuits:
            branch = CircuitBranch(circuit)
            if not branch.is_power_circuit:
                continue
            for warning in self.validator.validate(branch):
                if self.logger:
                    self.logger.warning(warning)
            branch.calculate_breaker_size()
            branch.calculate_hot_wire_size()
            branch.calculate_ground_wire_size()
            branch.calculate_conduit_size()
            branch.calculate_conduit_fill_percentage()
            branches.append(branch)

        tg = DB.TransactionGroup(self.doc, "Calculate Circuits")
        tg.Start()
        t = DB.Transaction(self.doc, "Write Shared Parameters")
        try:
            t.Start()
            for branch in branches:
                param_values = self.collect_shared_param_values(branch)
                self.update_circuit_parameters(branch.circuit, param_values)
                f, e = self.update_connected_elements(branch, param_values)
                total_fixtures += f
                total_equipment += e
            t.Commit()
            tg.Assimilate()
        except Exception as e:
            t.RollBack()
            tg.RollBack()
            if self.logger:
                self.logger.error("{}❌ Transaction failed: {}".format(branch.name, e))
            return

        return {
            'circuits': len(branches),
            'fixtures': total_fixtures,
            'equipment': total_equipment,
        }
