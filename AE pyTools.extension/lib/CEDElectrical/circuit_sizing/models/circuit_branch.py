from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, Optional


@dataclass
class CircuitSettings:
    """User-configurable sizing settings shared across calculations."""

    min_wire_size: str = "12"
    max_wire_size: str = "600"
    min_breaker_size: int = 20
    auto_calculate_breaker: bool = False
    min_conduit_size: str = '3/4"'
    max_conduit_fill: float = 0.36
    max_branch_voltage_drop: float = 0.03
    max_feeder_voltage_drop: float = 0.02
    wire_size_prefix: str = '#'
    conduit_size_suffix: str = 'C'

    def to_dict(self) -> Dict[str, object]:
        return {
            "min_wire_size": self.min_wire_size,
            "max_wire_size": self.max_wire_size,
            "min_breaker_size": self.min_breaker_size,
            "auto_calculate_breaker": self.auto_calculate_breaker,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, object]) -> "CircuitSettings":
        settings = cls()
        settings.min_wire_size = data.get("min_wire_size", settings.min_wire_size)
        settings.max_wire_size = data.get("max_wire_size", settings.max_wire_size)
        settings.min_breaker_size = data.get("min_breaker_size", settings.min_breaker_size)
        settings.auto_calculate_breaker = data.get(
            "auto_calculate_breaker", settings.auto_calculate_breaker
        )
        return settings


@dataclass
class CircuitOverrides:
    """User-entered override flags and values sourced from Revit."""

    auto_calculate: bool = False
    include_neutral: bool = False
    include_isolated_ground: bool = False

    breaker_override: Optional[float] = None
    wire_sets_override: Optional[int] = None
    wire_material_override: Optional[str] = None
    wire_temp_rating_override: Optional[str] = None
    wire_insulation_override: Optional[str] = None
    wire_hot_size_override: Optional[str] = None
    wire_neutral_size_override: Optional[str] = None
    wire_ground_size_override: Optional[str] = None
    conduit_type_override: Optional[str] = None
    conduit_size_override: Optional[str] = None


@dataclass
class CircuitBranchModel:
    """Pure-data representation of a Revit electrical circuit."""

    circuit_id: int
    panel: str
    circuit_number: str
    name: str
    branch_type: str

    rating: Optional[float] = None
    frame: Optional[float] = None
    length: Optional[float] = None
    voltage: Optional[float] = None
    apparent_power: Optional[float] = None
    apparent_current: Optional[float] = None
    circuit_load_current: Optional[float] = None
    poles: Optional[int] = None
    phase: int = 0
    power_factor: Optional[float] = None
    load_name: Optional[str] = None
    circuit_notes: str = ""

    wire_info: Dict[str, object] = field(default_factory=dict)
    overrides: CircuitOverrides = field(default_factory=CircuitOverrides)
    settings: CircuitSettings = field(default_factory=CircuitSettings)

    is_feeder: bool = False
    is_spare: bool = False
    is_space: bool = False
    is_transformer_primary: bool = False
    is_transformer_secondary: bool = False

    @property
    def max_voltage_drop(self) -> float:
        return (
            self.settings.max_feeder_voltage_drop
            if self.is_feeder
            else self.settings.max_branch_voltage_drop
        )

    @property
    def classification(self) -> str:
        """Return a friendly description of the circuit category."""
        if self.is_transformer_primary:
            return "TRANSFORMER_PRIMARY"
        if self.is_transformer_secondary:
            return "TRANSFORMER_SECONDARY"
        if self.is_feeder:
            return "FEEDER"
        if self.is_space:
            return "SPACE"
        if self.is_spare:
            return "SPARE"
        return "BRANCH"


@dataclass
class ConduitResult:
    size: Optional[str] = None
    fill: Optional[float] = None


@dataclass
class WireSizingResult:
    breaker_rating: Optional[float] = None
    hot_wire_size: Optional[str] = None
    wire_sets: Optional[int] = None
    hot_ampacity: Optional[float] = None
    ground_wire_size: Optional[str] = None
    voltage_drop: Optional[float] = None


@dataclass
class CircuitCalculationResult:
    wire: WireSizingResult = field(default_factory=WireSizingResult)
    conduit: ConduitResult = field(default_factory=ConduitResult)

    wire_material: Optional[str] = None
    wire_temp_rating: Optional[str] = None
    wire_insulation: Optional[str] = None

    neutral_included: bool = False
    isolated_ground_included: bool = False

    @property
    def number_of_sets(self) -> Optional[int]:
        return self.wire.wire_sets

    @property
    def circuit_base_ampacity(self) -> Optional[float]:
        return self.wire.hot_ampacity
