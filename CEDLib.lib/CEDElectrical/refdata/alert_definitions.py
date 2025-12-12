# -*- coding: utf-8 -*-
"""Central registry for Calculate Circuits alert definitions."""

from CEDElectrical.Model.alerts import AlertDefinition


ALERT_DEFINITIONS = {
    "overrides_invalid_circuit_property": AlertDefinition(
        "Overrides.InvalidCircuitProperty",
        "User-specified {property} '{override_value}' invalid; Using default {property}: {default_value}",
        group="Overrides",
    ),
    "overrides_invalid_equipment_ground": AlertDefinition(
        "Overrides.InvalidEquipmentGround",
        "User-specified equipment ground size '{override_value}' invalid; sizing per NEC Table 250.122",
        group="Overrides",
    ),
    "overrides_invalid_service_ground": AlertDefinition(
        "Overrides.InvalidServiceGround",
        "User-specified service ground size '{override_value}' invalid; sizing per NEC Table 250.102(c)",
        group="Overrides",
    ),
    "overrides_invalid_hot_wire": AlertDefinition(
        "Overrides.InvalidHotWire",
        "User-specified wire size '{override_value}' invalid; returning calculated value instead.",
        group="Overrides",
    ),
    "overrides_invalid_conduit": AlertDefinition(
        "Overrides.InvalidConduit",
        "User-specified conduit size '{override_value}' invalid; returning calculated value instead.",
        group="Overrides",
    ),
    "design_non_standard_ocp_rating": AlertDefinition(
        "Design.NonStandardOCPRating",
        "'{breaker_size}A' Circuit Breaker is a non-standard rating. Using next available size ({next_size}A) for calculations. Review your breaker ratings!",
        group="Design",
    ),
    "design_breaker_lug_size_limit_override": AlertDefinition(
        "Design.BreakerLugSizeLimitOverride",
        "User-specified Hot size {hot_size} exceeds recommended maximum for {breaker_size}A Breaker ({max_lug_size}).",
        group="Design",
    ),
    "design_breaker_lug_quantity_limit_override": AlertDefinition(
        "Design.BreakerLugQuantityLimitOverride",
        "User-specified wire sets ({wire_sets} sets) exceeds recommended maximum for {breaker_size}A Breaker ({max_lug_qty} sets).",
        group="Design",
    ),
    "calculation_breaker_lug_size_limit": AlertDefinition(
        "Calculations.BreakerLugSizeLimit",
        "Calculated Hot Size {hot_size} exceeds recommended maximum for breaker size {breaker_size}A. Review feeder lengths, connected loads, and breaker size!",
        group="Calculation",
    ),
    "calculation_breaker_lug_quantity_limit": AlertDefinition(
        "Calculations.BreakerLugQuantityLimit",
        "Calculated wire sets ({wire_sets} sets) exceeds recommended maximum for breaker size {breaker_size}A. Review feeder lengths, connected loads, and breaker size!",
        group="Calculation",
    ),
    "design_excessive_conduit_fill": AlertDefinition(
        "Design.ExcessiveConduitFill",
        "User-specified conduit size {conduit_size} exceeds max fill target ({conduit_fill_percentage}% < {max_fill_percentage}%).",
        group="Design",
    ),
    "design_undersized_wire_egc": AlertDefinition(
        "Design.UndersizedWireEGC",
        "User-specified ground size ({ground_size} {wire_material}) does not meet required size per NEC 250.122!",
        group="Design",
    ),
    "design_undersized_wire_service_ground": AlertDefinition(
        "Design.UndersizedWireServiceGround",
        "User-specified ground size ({ground_size} {wire_material}) does not meet required size per NEC 250.102!",
        group="Design",
    ),
    "design_excessive_volt_drop": AlertDefinition(
        "Design.ExcessiveVoltDrop",
        "User-specified wire ({wire_sets} set(s) x {wire_size}) fails volt drop check ({vd_percent}%). Review feeder lengths, connected loads and wire sizes!",
        group="Design",
    ),
    "design_insufficient_ampacity": AlertDefinition(
        "Design.InsufficientAmpacity",
        "User-specified wire ({wire_sets} set(s) x {wire_size}) fails ampacity check (Ampacity: {circuit_ampacity}A, Circuit Load: {circuit_load_current}A)",
        group="Design",
    ),
    "design_undersized_ocp": AlertDefinition(
        "Design.UndersizedOCP",
        "Circuit load current ({circuit_load_current}A) exceeds User-specified breaker rating ({breaker_rating}A)",
        group="Design",
    ),
}
