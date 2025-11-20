from CEDElectrical.circuit_sizing.domain.helpers import normalize_conduit_size, normalize_wire_size
from CEDElectrical.circuit_sizing.models.circuit_branch import CircuitBranchModel, CircuitCalculationResult
from CEDElectrical.refdata.conductor_area_table import CONDUCTOR_AREA_TABLE
from CEDElectrical.refdata.conduit_area_table import CONDUIT_AREA_TABLE, CONDUIT_SIZE_INDEX


class ConduitSizer:
    """Calculates conduit size and fill based on conductor quantities."""
    def __init__(self, model, results):
        self.model = model
        self.results = results

    def conduit_material(self):
        if self.model.overrides.conduit_type_override:
            return self._material_from_type(self.model.overrides.conduit_type_override)
        return self.model.wire_info.get("conduit_material_type")

    def _material_from_type(self, conduit_type):
        for material, type_dict in CONDUIT_AREA_TABLE.items():
            if conduit_type in type_dict:
                return material
        return None

    def calculate_conduit(self):
        insulation = self.results.wire_insulation or self.model.wire_info.get("wire_insulation")
        conduit_material = self.conduit_material()
        conduit_type = self.model.overrides.conduit_type_override or self.model.wire_info.get("conduit_type", "EMT")

        if not insulation or not conduit_material or not conduit_type:
            return None

        sizes_and_qtys = [
            (normalize_wire_size(self._formatted_hot_wire(), self.model.settings.wire_size_prefix), self.hot_wire_quantity()),
            (normalize_wire_size(self._formatted_neutral_wire(), self.model.settings.wire_size_prefix), self.neutral_wire_quantity()),
            (normalize_wire_size(self._formatted_ground_wire(), self.model.settings.wire_size_prefix), self.ground_wire_quantity()),
            (normalize_wire_size(self._formatted_isolated_ground_wire(), self.model.settings.wire_size_prefix), self.isolated_ground_quantity()),
        ]

        total_area = sum(
            CONDUCTOR_AREA_TABLE[size]["area"][insulation] * qty
            for size, qty in sizes_and_qtys
            if size and qty and size in CONDUCTOR_AREA_TABLE and insulation in CONDUCTOR_AREA_TABLE[size]["area"]
        )

        conduit_table = CONDUIT_AREA_TABLE.get(conduit_material, {}).get(conduit_type, {})
        if not conduit_table:
            return None

        enum = CONDUIT_SIZE_INDEX
        if self.model.settings.min_conduit_size not in enum:
            return None
        min_index = enum.index(self.model.settings.min_conduit_size)

        for size in enum[min_index:]:
            area = conduit_table.get(size)
            if area and area * self.model.settings.max_conduit_fill >= total_area:
                self.results.conduit.size = size
                self.results.conduit.fill = round(total_area / area, 5)
                return size
        return None

    def calculate_fill(self):
        conduit_raw = self.model.overrides.conduit_size_override or self.results.conduit.size
        conduit_size = normalize_conduit_size(conduit_raw, self.model.settings.conduit_size_suffix)
        insulation = self.results.wire_insulation or self.model.wire_info.get("wire_insulation")
        conduit_material = self.conduit_material()
        conduit_type = self.model.overrides.conduit_type_override or self.model.wire_info.get("conduit_type")

        if not conduit_size or not conduit_material or not conduit_type or not insulation:
            return None

        conduit_area = CONDUIT_AREA_TABLE.get(conduit_material, {}).get(conduit_type, {}).get(conduit_size)
        if not conduit_area:
            return None

        sizes_and_qtys = [
            (normalize_wire_size(self._formatted_hot_wire(), self.model.settings.wire_size_prefix), self.hot_wire_quantity()),
            (normalize_wire_size(self._formatted_neutral_wire(), self.model.settings.wire_size_prefix), self.neutral_wire_quantity()),
            (normalize_wire_size(self._formatted_ground_wire(), self.model.settings.wire_size_prefix), self.ground_wire_quantity()),
            (normalize_wire_size(self._formatted_isolated_ground_wire(), self.model.settings.wire_size_prefix), self.isolated_ground_quantity()),
        ]

        total_area = sum(
            CONDUCTOR_AREA_TABLE[size]["area"][insulation] * qty
            for size, qty in sizes_and_qtys
            if size and qty and size in CONDUCTOR_AREA_TABLE and insulation in CONDUCTOR_AREA_TABLE[size]["area"]
        )

        self.results.conduit.fill = round(total_area / conduit_area, 5)
        return self.results.conduit.fill

    # --- Quantities/formatting helpers ---
    def hot_wire_quantity(self):
        return self.model.poles or 0

    def neutral_wire_quantity(self):
        return 1 if self.results.neutral_included else 0

    def ground_wire_quantity(self):
        return 1 if self.model.branch_type != "SPACE" else 0

    def isolated_ground_quantity(self):
        return 1 if self.results.isolated_ground_included else 0

    def _formatted_hot_wire(self):
        if self.model.overrides.wire_hot_size_override and self.model.overrides.auto_calculate:
            return self.model.overrides.wire_hot_size_override
        if not self.results.wire.hot_wire_size:
            return None
        if self.model.settings.wire_size_prefix:
            return "{}{}".format(self.model.settings.wire_size_prefix, self.results.wire.hot_wire_size)
        return self.results.wire.hot_wire_size

    def _formatted_neutral_wire(self):
        if not self.results.neutral_included:
            return None
        if self.model.overrides.wire_neutral_size_override and self.model.overrides.auto_calculate:
            return self.model.overrides.wire_neutral_size_override
        return self._formatted_hot_wire()

    def _formatted_ground_wire(self):
        if self.model.overrides.wire_ground_size_override and self.model.overrides.auto_calculate:
            return self.model.overrides.wire_ground_size_override
        if not self.results.wire.ground_wire_size:
            return None
        if self.model.settings.wire_size_prefix:
            return "{}{}".format(self.model.settings.wire_size_prefix, self.results.wire.ground_wire_size)
        return self.results.wire.ground_wire_size

    def _formatted_isolated_ground_wire(self):
        if not self.results.isolated_ground_included:
            return None
        return self._formatted_ground_wire()
