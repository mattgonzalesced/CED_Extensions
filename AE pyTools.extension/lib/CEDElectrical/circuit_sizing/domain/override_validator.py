from __future__ import annotations

from typing import Optional

from CEDElectrical.circuit_sizing.domain.helpers import normalize_conduit_size, normalize_wire_size
from CEDElectrical.circuit_sizing.models.circuit_branch import CircuitBranchModel, CircuitOverrides


class OverrideValidator:
    """Cleans and normalizes user-entered overrides before calculations."""

    def __init__(self, model: CircuitBranchModel):
        self.model = model
        self.overrides: CircuitOverrides = model.overrides

    def cleaned_overrides(self) -> CircuitOverrides:
        settings = self.model.settings
        cleaned = CircuitOverrides(
            auto_calculate=self.overrides.auto_calculate,
            include_neutral=self.overrides.include_neutral,
            include_isolated_ground=self.overrides.include_isolated_ground,
        )

        cleaned.breaker_override = self._safe_float(self.overrides.breaker_override)
        cleaned.wire_sets_override = self._safe_int(self.overrides.wire_sets_override)

        cleaned.wire_material_override = self._safe_str(self.overrides.wire_material_override)
        cleaned.wire_temp_rating_override = self._safe_str(
            self.overrides.wire_temp_rating_override
        )
        cleaned.wire_insulation_override = self._safe_str(
            self.overrides.wire_insulation_override
        )

        cleaned.wire_hot_size_override = normalize_wire_size(
            self.overrides.wire_hot_size_override, settings.wire_size_prefix
        )
        cleaned.wire_neutral_size_override = normalize_wire_size(
            self.overrides.wire_neutral_size_override, settings.wire_size_prefix
        )
        cleaned.wire_ground_size_override = normalize_wire_size(
            self.overrides.wire_ground_size_override, settings.wire_size_prefix
        )

        cleaned.conduit_type_override = self._safe_str(
            self.overrides.conduit_type_override
        )
        cleaned.conduit_size_override = normalize_conduit_size(
            self.overrides.conduit_size_override, settings.conduit_size_suffix
        )
        return cleaned

    @staticmethod
    def _safe_str(value: Optional[str]) -> Optional[str]:
        return str(value).strip() if value not in (None, "") else None

    @staticmethod
    def _safe_float(value: Optional[float]) -> Optional[float]:
        try:
            return float(value) if value not in (None, "") else None
        except Exception:
            return None

    @staticmethod
    def _safe_int(value: Optional[int]) -> Optional[int]:
        try:
            return int(value) if value not in (None, "") else None
        except Exception:
            return None
