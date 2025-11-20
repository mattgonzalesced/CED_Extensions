from __future__ import annotations

from typing import Optional, Tuple

from CEDElectrical.circuit_sizing.domain.helpers import normalize_wire_size
from CEDElectrical.circuit_sizing.domain.voltage_drop import VoltageDropCalculator
from CEDElectrical.circuit_sizing.models.circuit_branch import CircuitBranchModel, CircuitCalculationResult
from CEDElectrical.refdata.ampacity_table import WIRE_AMPACITY_TABLE
from CEDElectrical.refdata.conductor_area_table import CONDUCTOR_AREA_TABLE
from CEDElectrical.refdata.egc_table import EGC_TABLE
from CEDElectrical.refdata.standard_ocp_table import BREAKER_FRAME_SWITCH_TABLE


class WireSizer:
    """Determines breaker, conductor sizes, and ampacity for a circuit model."""
    def __init__(self, model: CircuitBranchModel):
        self.model = model

    def calculate_breaker(self) -> Optional[float]:
        """Return a breaker size based on load and minimum settings."""
        if self.model.settings.auto_calculate_breaker:
            amps = self.model.apparent_current
        else:
            return self.model.rating

        if not amps:
            return None

        amps *= 1.25
        if amps < self.model.settings.min_breaker_size:
            amps = self.model.settings.min_breaker_size

        for breaker in sorted(BREAKER_FRAME_SWITCH_TABLE.keys()):
            if breaker >= amps:
                return breaker
        return None

    def size_hot_conductor(
        self, overrides: CircuitCalculationResult, cleaned_hot_override: Optional[str]
    ) -> Tuple[Optional[str], Optional[int], Optional[float]]:
        """Size the hot conductor based on overrides and ampacity rules."""
        rating = overrides.wire.breaker_rating
        wire_info = self.model.wire_info
        settings = self.model.settings

        if rating is None:
            return None, None, None

        if self.model.overrides.auto_calculate and cleaned_hot_override:
            material = overrides.wire_material or wire_info.get("wire_material")
            temp = int(str(overrides.wire_temp_rating or wire_info.get("wire_temperature_rating", "75")).replace("C", "").strip())
            sets = overrides.wire_sets or self.model.overrides.wire_sets_override or 1
            wire_set = WIRE_AMPACITY_TABLE.get(material, {}).get(temp, [])
            for wire, ampacity in wire_set:
                if wire == cleaned_hot_override:
                    return wire, sets, ampacity * sets
            return cleaned_hot_override, sets, None

        try:
            temp = int(str(wire_info.get("wire_temperature_rating", "75")).replace("C", "").strip())
        except Exception:
            return None, None, None

        material = overrides.wire_material or wire_info.get("wire_material", "CU")
        base_wire = wire_info.get("wire_hot_size")
        base_sets = wire_info.get("number_of_parallel_sets") or 1
        max_size = wire_info.get("max_lug_size")
        max_sets = wire_info.get("max_lug_qty", 1)

        wire_set = WIRE_AMPACITY_TABLE.get(material, {}).get(temp, [])
        sets = base_sets or 1
        start_index = 0
        for i, (wire, _) in enumerate(wire_set):
            if wire == base_wire:
                start_index = i
                break

        while sets <= max_sets:
            reached_max_size = False
            for wire, ampacity in wire_set[start_index:]:
                total_ampacity = ampacity * sets
                if not self._is_ampacity_acceptable(rating, total_ampacity, self.model.circuit_load_current):
                    continue

                vd = VoltageDropCalculator.calculate(self.model, wire, sets)
                if vd is None or vd <= self.model.max_voltage_drop:
                    return wire, sets, total_ampacity

                if wire == max_size:
                    return wire, sets, total_ampacity

            if reached_max_size:
                break
            sets += 1

        return None, None, None

    def size_ground_conductor(
        self, overrides: CircuitCalculationResult, cleaned_ground_override: Optional[str]
    ) -> Optional[str]:
        rating = overrides.wire.breaker_rating
        wire_info = self.model.wire_info
        material = overrides.wire_material or wire_info.get("wire_material", "CU")

        if rating is None:
            return None

        if cleaned_ground_override:
            return cleaned_ground_override

        base_ground = wire_info.get("wire_ground_size")
        base_hot = wire_info.get("wire_hot_size")
        base_sets = wire_info.get("number_of_parallel_sets", 1)

        calc_hot = overrides.wire.hot_wire_size
        calc_sets = overrides.wire.wire_sets

        if not base_ground:
            egc_list = EGC_TABLE.get(material)
            if egc_list:
                for threshold, size in egc_list:
                    if rating <= threshold:
                        return size
                return egc_list[-1][1]
            return None

        if not (base_ground and base_hot and calc_hot and calc_sets):
            return None

        base_hot_cmil = CONDUCTOR_AREA_TABLE.get(base_hot, {}).get("cmil")
        calc_hot_cmil = CONDUCTOR_AREA_TABLE.get(calc_hot, {}).get("cmil")
        base_ground_cmil = CONDUCTOR_AREA_TABLE.get(base_ground, {}).get("cmil")

        if not all([base_hot_cmil, calc_hot_cmil, base_ground_cmil]):
            return None

        total_base_hot_cmil = base_sets * base_hot_cmil
        total_calc_hot_cmil = calc_sets * calc_hot_cmil
        new_ground_cmil = base_ground_cmil * (float(total_calc_hot_cmil) / total_base_hot_cmil)

        for wire, data in sorted(CONDUCTOR_AREA_TABLE.items(), key=lambda x: x[1]["cmil"]):
            if data["cmil"] >= new_ground_cmil:
                return wire
        return None

    @staticmethod
    def _is_ampacity_acceptable(breaker_rating: float, ampacity: float, circuit_amps: Optional[float]) -> bool:
        if circuit_amps is not None and ampacity < circuit_amps:
            return False
        if ampacity >= breaker_rating:
            return True
        if breaker_rating > 800:
            return False

        for std_rating in sorted(BREAKER_FRAME_SWITCH_TABLE.keys()):
            if std_rating >= ampacity:
                return std_rating >= breaker_rating
        return False

    def evaluate(self, cleaned_overrides) -> CircuitCalculationResult:
        result = CircuitCalculationResult()
        result.wire_material = cleaned_overrides.wire_material_override or self.model.wire_info.get("wire_material")
        result.wire_temp_rating = cleaned_overrides.wire_temp_rating_override or self.model.wire_info.get("wire_temperature_rating")
        result.wire_insulation = cleaned_overrides.wire_insulation_override or self.model.wire_info.get("wire_insulation")

        if cleaned_overrides.auto_calculate and cleaned_overrides.breaker_override:
            result.wire.breaker_rating = cleaned_overrides.breaker_override
        else:
            result.wire.breaker_rating = self.calculate_breaker()

        hot_wire, wire_sets, ampacity = self.size_hot_conductor(result, cleaned_overrides.wire_hot_size_override)
        result.wire.hot_wire_size = hot_wire
        result.wire.wire_sets = cleaned_overrides.wire_sets_override or wire_sets
        result.wire.hot_ampacity = ampacity

        ground_wire = self.size_ground_conductor(result, cleaned_overrides.wire_ground_size_override)
        result.wire.ground_wire_size = ground_wire

        result.neutral_included = cleaned_overrides.include_neutral or (self.model.poles == 1)
        result.isolated_ground_included = cleaned_overrides.include_isolated_ground

        result.wire.voltage_drop = VoltageDropCalculator.calculate(
            self.model, result.wire.hot_wire_size, result.wire.wire_sets or 1
        )
        return result
