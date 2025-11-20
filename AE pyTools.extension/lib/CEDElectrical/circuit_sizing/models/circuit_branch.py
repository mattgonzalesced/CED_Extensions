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
        return {
            "min_wire_size": self.min_wire_size,
            "max_wire_size": self.max_wire_size,
            "min_breaker_size": self.min_breaker_size,
            "auto_calculate_breaker": self.auto_calculate_breaker,
        }

    @classmethod
    def from_dict(cls, data):
        """Create a settings object from a dictionary."""
        settings = cls()
        settings.min_wire_size = data.get("min_wire_size", settings.min_wire_size)
        settings.max_wire_size = data.get("max_wire_size", settings.max_wire_size)
        settings.min_breaker_size = data.get("min_breaker_size", settings.min_breaker_size)
        settings.auto_calculate_breaker = data.get(
            "auto_calculate_breaker", settings.auto_calculate_breaker
        )
        return settings


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
        return self.wire.hot_ampacity
