# -*- coding: utf-8 -*-
"""Adapters that push calculation results back into Revit."""

import Autodesk.Revit.DB as DB
import Autodesk.Revit.DB.Electrical as DBE


class RevitCircuitWriter(object):
    def __init__(self, doc):
        self.doc = doc

    def _set_param(self, element, name, value):
        if value is None:
            return
        param = element.LookupParameter(name)
        if not param:
            return
        try:
            if param.StorageType == DB.StorageType.String:
                param.Set(str(value))
            elif param.StorageType == DB.StorageType.Integer:
                param.Set(int(value))
            elif param.StorageType == DB.StorageType.Double:
                param.Set(float(value))
        except Exception:
            pass

    def _collect_values(self, result):
        model = result.model
        return {
            'CKT_Circuit Type_CEDT': model.branch_type,
            'CKT_Panel_CEDT': model.panel,
            'CKT_Circuit Number_CEDT': model.circuit_number,
            'CKT_Load Name_CEDT': model.load_name,
            'CKT_Rating_CED': result.breaker_rating,
            'CKT_Frame_CED': result.frame,
            'CKT_Length_CED': result.length,
            'CKT_Schedule Notes_CEDT': result.circuit_notes,
            'Voltage Drop Percentage_CED': result.voltage_drop_percentage,
            'CKT_Wire Hot Size_CEDT': result.hot_wire_size,
            'CKT_Number of Wires_CED': result.number_of_wires,
            'CKT_Number of Sets_CED': result.number_of_sets,
            'CKT_Wire Hot Quantity_CED': result.hot_wire_quantity,
            'CKT_Wire Ground Size_CEDT': result.ground_wire_size,
            'CKT_Wire Ground Quantity_CED': result.ground_wire_quantity,
            'CKT_Wire Neutral Size_CEDT': result.neutral_wire_size,
            'CKT_Wire Neutral Quantity_CED': result.neutral_wire_quantity,
            'CKT_Wire Isolated Ground Size_CEDT': result.isolated_ground_wire_size,
            'CKT_Wire Isolated Ground Quantity_CED': result.isolated_ground_wire_quantity,
            'Wire Material_CEDT': result.wire_material,
            'Wire Temparature Rating_CEDT': result.wire_temp_rating,
            'Wire Insulation_CEDT': result.wire_insulation,
            'Conduit Size_CEDT': result.conduit_size,
            'Conduit Type_CEDT': result.conduit_type,
            'Conduit Fill Percentage_CED': result.conduit_fill_percentage,
            'Wire Size_CEDT': result.wire_size_callout,
            'Conduit and Wire Size_CEDT': result.conduit_and_wire_size,
            'Circuit Load Current_CED': result.circuit_load_current,
            'Circuit Ampacity_CED': result.circuit_base_ampacity,
        }

    def write_circuit(self, circuit, result):
        values = self._collect_values(result)
        for name, value in values.items():
            self._set_param(circuit, name, value)

    def write_connected(self, circuit, result):
        values = self._collect_values(result)
        fixture_count = 0
        equipment_count = 0

        for el in circuit.Elements:
            if not isinstance(el, DBE.FamilyInstance):
                continue
            cat = el.Category
            if not cat:
                continue
            cat_id = cat.Id
            is_fixture = cat_id == DB.ElementId(DB.BuiltInCategory.OST_ElectricalFixtures)
            is_equipment = cat_id == DB.ElementId(DB.BuiltInCategory.OST_ElectricalEquipment)
            if not (is_fixture or is_equipment):
                continue
            for name, value in values.items():
                self._set_param(el, name, value)
            if is_fixture:
                fixture_count += 1
            if is_equipment:
                equipment_count += 1
        return fixture_count, equipment_count
