<<<<<<< ours
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
=======
"""Data-only classes supporting circuit sizing logic."""


class CircuitSettings(object):
    """User-configurable sizing settings shared across calculations."""

    def __init__(
        self,
        min_wire_size="12",
        max_wire_size="600",
        min_breaker_size=20,
        auto_calculate_breaker=False,
        min_conduit_size='3/4"',
        max_conduit_fill=0.36,
        max_branch_voltage_drop=0.03,
        max_feeder_voltage_drop=0.02,
        wire_size_prefix="#",
        conduit_size_suffix="C",
    ):
        self.min_wire_size = min_wire_size
        self.max_wire_size = max_wire_size
        self.min_breaker_size = min_breaker_size
        self.auto_calculate_breaker = auto_calculate_breaker
        self.min_conduit_size = min_conduit_size
        self.max_conduit_fill = max_conduit_fill
        self.max_branch_voltage_drop = max_branch_voltage_drop
        self.max_feeder_voltage_drop = max_feeder_voltage_drop
        self.wire_size_prefix = wire_size_prefix
        self.conduit_size_suffix = conduit_size_suffix

    def to_dict(self):
        """Serialize minimal settings to a dictionary."""
>>>>>>> theirs
        return {
            "min_wire_size": self.min_wire_size,
            "max_wire_size": self.max_wire_size,
            "min_breaker_size": self.min_breaker_size,
            "auto_calculate_breaker": self.auto_calculate_breaker,
        }

    @classmethod
<<<<<<< ours
    def from_dict(cls, data: Dict[str, object]) -> "CircuitSettings":
=======
    def from_dict(cls, data):
        """Create a settings object from a dictionary."""
>>>>>>> theirs
        settings = cls()
        settings.min_wire_size = data.get("min_wire_size", settings.min_wire_size)
        settings.max_wire_size = data.get("max_wire_size", settings.max_wire_size)
        settings.min_breaker_size = data.get("min_breaker_size", settings.min_breaker_size)
        settings.auto_calculate_breaker = data.get(
            "auto_calculate_breaker", settings.auto_calculate_breaker
        )
        return settings


<<<<<<< ours
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
=======
class CircuitOverrides(object):
    """User-entered override flags and values sourced from Revit."""

    def __init__(self):
        self.auto_calculate = False
        self.include_neutral = False
        self.include_isolated_ground = False

        self.breaker_override = None
        self.wire_sets_override = None
        self.wire_material_override = None
        self.wire_temp_rating_override = None
        self.wire_insulation_override = None
        self.wire_hot_size_override = None
        self.wire_neutral_size_override = None
        self.wire_ground_size_override = None
        self.conduit_type_override = None
        self.conduit_size_override = None


class CircuitBranchModel(object):
    """Pure-data representation of a Revit electrical circuit."""

    def __init__(
        self,
        circuit_id,
        panel,
        circuit_number,
        name,
        branch_type,
        rating=None,
        frame=None,
        length=None,
        voltage=None,
        apparent_power=None,
        apparent_current=None,
        circuit_load_current=None,
        poles=None,
        phase=0,
        power_factor=None,
        load_name=None,
        circuit_notes="",
        wire_info=None,
        overrides=None,
        settings=None,
        is_feeder=False,
        is_spare=False,
        is_space=False,
        is_transformer_primary=False,
        is_transformer_secondary=False,
    ):
        self.circuit_id = circuit_id
        self.panel = panel
        self.circuit_number = circuit_number
        self.name = name
        self.branch_type = branch_type

        self.rating = rating
        self.frame = frame
        self.length = length
        self.voltage = voltage
        self.apparent_power = apparent_power
        self.apparent_current = apparent_current
        self.circuit_load_current = circuit_load_current
        self.poles = poles
        self.phase = phase
        self.power_factor = power_factor
        self.load_name = load_name
        self.circuit_notes = circuit_notes or ""

        self.wire_info = wire_info or {}
        self.overrides = overrides or CircuitOverrides()
        self.settings = settings or CircuitSettings()

        self.is_feeder = is_feeder
        self.is_spare = is_spare
        self.is_space = is_space
        self.is_transformer_primary = is_transformer_primary
        self.is_transformer_secondary = is_transformer_secondary

    @property
    def max_voltage_drop(self):
        """Maximum allowed voltage drop using feeder/branch thresholds."""
        if self.is_feeder:
            return self.settings.max_feeder_voltage_drop
        return self.settings.max_branch_voltage_drop

    @property
    def classification(self):
>>>>>>> theirs
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


<<<<<<< ours
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
=======
class ConduitResult(object):
    """Container for conduit sizing output."""

    def __init__(self):
        self.size = None
        self.fill = None


class WireSizingResult(object):
    """Container for wire sizing output."""

    def __init__(self):
        self.breaker_rating = None
        self.hot_wire_size = None
        self.wire_sets = None
        self.hot_ampacity = None
        self.ground_wire_size = None
        self.voltage_drop = None


class CircuitCalculationResult(object):
    """Composite result structure used by evaluators and writers."""

    def __init__(self):
        self.wire = WireSizingResult()
        self.conduit = ConduitResult()

        self.wire_material = None
        self.wire_temp_rating = None
        self.wire_insulation = None

        self.neutral_included = False
        self.isolated_ground_included = False

    @property
    def number_of_sets(self):
        """Return the total number of wire sets for convenience."""
        return self.wire.wire_sets

    @property
    def circuit_base_ampacity(self):
        """Return the hot conductor ampacity for writing back to Revit."""
>>>>>>> theirs
        return self.wire.hot_ampacity
