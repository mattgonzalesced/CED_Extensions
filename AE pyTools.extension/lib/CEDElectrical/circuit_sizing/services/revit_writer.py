<<<<<<< ours
from __future__ import annotations

from typing import Dict, Tuple

=======
>>>>>>> theirs
import Autodesk.Revit.DB as DB
from pyrevit import script

from CEDElectrical.circuit_sizing.domain.helpers import normalize_conduit_size, normalize_wire_size
from CEDElectrical.circuit_sizing.models.circuit_branch import CircuitBranchModel, CircuitCalculationResult

logger = script.get_logger()


class RevitCircuitWriter:
    """Writes circuit calculation results back into Revit parameters."""
    @staticmethod
<<<<<<< ours
    def collect_param_values(model: CircuitBranchModel, results: CircuitCalculationResult) -> Dict[str, object]:
=======
    def collect_param_values(model, results):
>>>>>>> theirs
        hot_qty = model.poles or 0
        neutral_qty = 1 if results.neutral_included else 0
        ground_qty = 1 if model.branch_type != "SPACE" else 0
        isolated_ground_qty = 1 if results.isolated_ground_included else 0

        hot_size = RevitCircuitWriter._format_wire(results.wire.hot_wire_size, model.settings.wire_size_prefix)
        neutral_size = (
            RevitCircuitWriter._format_wire(results.wire.hot_wire_size, model.settings.wire_size_prefix)
            if neutral_qty
            else ""
        )

        ground_size = RevitCircuitWriter._format_wire(results.wire.ground_wire_size, model.settings.wire_size_prefix)
        iso_ground_size = ground_size if isolated_ground_qty else ""

        conduit_size = RevitCircuitWriter._format_conduit(
            normalize_conduit_size(model.overrides.conduit_size_override or results.conduit.size, model.settings.conduit_size_suffix),
            model.settings.conduit_size_suffix,
        )
        conduit_type = model.overrides.conduit_type_override or model.wire_info.get("conduit_type")

        wire_set_string = RevitCircuitWriter._wire_set_string(
            hot_qty,
            hot_size,
            neutral_qty,
            neutral_size,
            ground_qty,
            ground_size,
            isolated_ground_qty,
            iso_ground_size,
            model.wire_info.get("wire_material"),
            model.settings.wire_size_prefix,
        )

        return {
            'CKT_Circuit Type_CEDT': model.branch_type,
            'CKT_Panel_CEDT': model.panel,
            'CKT_Circuit Number_CEDT': model.circuit_number,
            'CKT_Load Name_CEDT': model.load_name,
            'CKT_Rating_CED': model.rating,
            'CKT_Frame_CED': model.frame,
            'CKT_Length_CED': model.length,
            'CKT_Schedule Notes_CEDT': model.circuit_notes,
            'Voltage Drop Percentage_CED': results.wire.voltage_drop,
            'CKT_Wire Hot Size_CEDT': hot_size,
            'CKT_Number of Wires_CED': hot_qty + neutral_qty,
            'CKT_Number of Sets_CED': results.number_of_sets,
            'CKT_Wire Hot Quantity_CED': hot_qty,
            'CKT_Wire Ground Size_CEDT': ground_size,
            'CKT_Wire Ground Quantity_CED': ground_qty,
            'CKT_Wire Neutral Size_CEDT': neutral_size,
            'CKT_Wire Neutral Quantity_CED': neutral_qty,
            'CKT_Wire Isolated Ground Size_CEDT': iso_ground_size,
            'CKT_Wire Isolated Ground Quantity_CED': isolated_ground_qty,
            'Wire Material_CEDT': results.wire_material,
            'Wire Temparature Rating_CEDT': results.wire_temp_rating,
            'Wire Insulation_CEDT': results.wire_insulation,
            'Conduit Size_CEDT': conduit_size,
            'Conduit Type_CEDT': conduit_type,
            'Conduit Fill Percentage_CED': results.conduit.fill,
            'Wire Size_CEDT': RevitCircuitWriter._wire_size_callout(results.number_of_sets, wire_set_string),
            'Conduit and Wire Size_CEDT': RevitCircuitWriter._conduit_and_wire_size(results.number_of_sets, conduit_size, conduit_type, wire_set_string),
            'Circuit Load Current_CED': model.circuit_load_current,
            'Circuit Ampacity_CED': results.circuit_base_ampacity,
        }

    @staticmethod
<<<<<<< ours
    def _format_wire(size: object, prefix: str) -> str:
=======
    def _format_wire(size, prefix):
>>>>>>> theirs
        if size is None:
            return ""
        if prefix:
            return "{}{}".format(prefix, size)
        return str(size)

    @staticmethod
<<<<<<< ours
    def _format_conduit(size: object, suffix: str) -> str:
=======
    def _format_conduit(size, suffix):
>>>>>>> theirs
        if size is None:
            return ""
        if suffix:
            return "{}{}".format(size, suffix)
        return str(size)

    @staticmethod
    def _wire_set_string(hot_qty, hot_size, neutral_qty, neutral_size, ground_qty, ground_size, ig_qty, ig_size, material, prefix):
        parts = []
        wire_prefix = prefix or ""
        if hot_qty:
            parts.append("{}{}{}".format(hot_qty, wire_prefix, normalize_wire_size(hot_size, prefix)))
        if neutral_qty:
            parts.append("{}{}{}N".format(neutral_qty, wire_prefix, normalize_wire_size(neutral_size, prefix)))
        if ground_qty:
            parts.append("{}{}{}G".format(ground_qty, wire_prefix, normalize_wire_size(ground_size, prefix)))
        if ig_qty:
            parts.append("{}{}{}IG".format(ig_qty, wire_prefix, normalize_wire_size(ig_size, prefix)))
        suffix = material if material and material != "CU" else ""
        return "{} {}".format(" + ".join(parts), suffix).strip()

    @staticmethod
<<<<<<< ours
    def _wire_size_callout(sets: int, wire_set_string: str) -> str:
=======
    def _wire_size_callout(sets, wire_set_string):
>>>>>>> theirs
        if not wire_set_string:
            return ""
        sets = sets or 1
        if sets > 1:
            return "({}) {}".format(sets, wire_set_string)
        return wire_set_string

    @staticmethod
<<<<<<< ours
    def _conduit_and_wire_size(sets: int, conduit: str, conduit_type: str, wire_callout: str) -> str:
=======
    def _conduit_and_wire_size(sets, conduit, conduit_type, wire_callout):
>>>>>>> theirs
        if not conduit:
            return ""
        prefix = "{} SETS - ".format(sets) if (sets or 1) > 1 else ""
        if wire_callout:
            return "{}{}-({})".format(prefix, conduit, wire_callout)
        return "{}{} ({})".format(prefix, conduit, conduit_type)

    @staticmethod
<<<<<<< ours
    def update_circuit_parameters(circuit, param_values: Dict[str, object]):
=======
    def update_circuit_parameters(circuit, param_values):
>>>>>>> theirs
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
                logger.debug("Failed to write '{}' to circuit {}: {}".format(param_name, circuit.Id, e))

    @staticmethod
<<<<<<< ours
    def update_connected_elements(circuit, param_values: Dict[str, object]) -> Tuple[int, int]:
=======
    def update_connected_elements(circuit, param_values):
>>>>>>> theirs
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
                    logger.debug("Failed to write '{}' to element {}: {}".format(param_name, el.Id, e))

            if is_fixture:
                fixture_count += 1
            elif is_equipment:
                equipment_count += 1

        return fixture_count, equipment_count
