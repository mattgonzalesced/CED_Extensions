# -*- coding: utf-8 -*-
"""Revit-backed circuit writer adapter."""

from pyrevit import DB


class RevitCircuitWriter(object):
    """Writes calculated circuit and downstream parameter values."""

    def write_circuit_parameters(self, circuit, param_values):
        """Write calculated parameter map to a circuit element."""
        for param_name, value in param_values.items():
            param = circuit.LookupParameter(param_name)
            if not param:
                continue
            try:
                st = param.StorageType
                if value is None:
                    if st == DB.StorageType.String:
                        param.Set('')
                    elif st == DB.StorageType.Integer:
                        param.Set(0)
                    elif st == DB.StorageType.Double:
                        param.Set(0.0)
                    elif st == DB.StorageType.ElementId:
                        param.Set(DB.ElementId.InvalidElementId)
                    continue

                if st == DB.StorageType.String:
                    param.Set(str(value))
                elif st == DB.StorageType.Integer:
                    param.Set(int(value))
                elif st == DB.StorageType.Double:
                    param.Set(float(value))
                elif st == DB.StorageType.ElementId and isinstance(value, DB.ElementId):
                    param.Set(value)
            except Exception:
                continue

    def write_connected_elements(self, branch, param_values, settings, locked_ids=None):
        """Write calculated values to connected fixtures/equipment."""
        circuit = branch.circuit
        fixture_count = 0
        equipment_count = 0
        locked_ids = locked_ids or set()

        write_fixtures = getattr(settings, 'write_fixture_results', False)
        write_equipment = getattr(settings, 'write_equipment_results', False)
        if not (write_fixtures or write_equipment):
            return fixture_count, equipment_count

        for el in circuit.Elements:
            if not isinstance(el, DB.FamilyInstance):
                continue
            if el.Id in locked_ids:
                continue

            cat = el.Category
            if not cat:
                continue

            cat_id = cat.Id
            is_fixture = cat_id == DB.ElementId(DB.BuiltInCategory.OST_ElectricalFixtures)
            is_equipment = cat_id == DB.ElementId(DB.BuiltInCategory.OST_ElectricalEquipment)

            if not (is_fixture or is_equipment):
                continue
            if is_fixture and not write_fixtures:
                continue
            if is_equipment and not write_equipment:
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
                except Exception:
                    continue

            if is_fixture:
                fixture_count += 1
            elif is_equipment:
                equipment_count += 1

        return fixture_count, equipment_count
