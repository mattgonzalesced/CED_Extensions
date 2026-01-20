# -*- coding: utf-8 -*-
"""Central registry for Calculate Circuits alert definitions."""

from CEDElectrical.Model.alerts import AlertDefinition


ALERT_DEFINITIONS = {
    "overrides_invalid_circuit_property": AlertDefinition(
        "Overrides.InvalidCircuitProperty",
        "User-specified {property} '{override_value}' invalid; Using default {property}: {default_value}",
        group="Overrides",
        severity="NONE",
    ),
    "overrides_invalid_equipment_ground": AlertDefinition(
        "Overrides.InvalidEquipmentGround",
        "User-specified equipment ground size '{override_value}' invalid; sizing per NEC Table 250.122",
        group="Overrides",
        severity="NONE",
    ),
    "overrides_invalid_service_ground": AlertDefinition(
        "Overrides.InvalidServiceGround",
        "User-specified service ground size '{override_value}' invalid; sizing per NEC Table 250.102(c)",
        group="Overrides",
        severity="NONE",
    ),
    "overrides_invalid_hot_wire": AlertDefinition(
        "Overrides.InvalidHotWire",
        "User-specified wire size '{override_value}' invalid; returning calculated value instead.",
        group="Overrides",
        severity="NONE",
    ),
    "overrides_invalid_conduit": AlertDefinition(
        "Overrides.InvalidConduit",
        "User-specified conduit size '{override_value}' invalid; returning calculated value instead.",
        group="Overrides",
        severity="NONE",
    ),
    "design_non_standard_ocp_rating": AlertDefinition(
        "Design.NonStandardOCPRating",
        "'{breaker_size}A' Circuit Breaker is a non-standard rating. Using next available size ({next_size}A) for calculations. Review your breaker ratings!",
        group="Design",
        severity="MEDIUM",
    ),
    "design_breaker_lug_size_limit_override": AlertDefinition(
        "Design.BreakerLugSizeLimitOverride",
        "User-specified Hot size {hot_size} exceeds recommended maximum for {breaker_size}A Breaker ({max_lug_size}).",
        group="Design",
        severity="MEDIUM",
    ),
    "design_breaker_lug_quantity_limit_override": AlertDefinition(
        "Design.BreakerLugQuantityLimitOverride",
        "User-specified wire sets ({wire_sets} sets) exceeds recommended maximum for {breaker_size}A Breaker ({max_lug_qty} sets).",
        group="Design",
        severity="MEDIUM",
    ),
    "calculation_breaker_lug_size_limit": AlertDefinition(
        "Calculations.BreakerLugSizeLimit",
        "Calculated Hot Size #{hot_size} exceeds recommended maximum for breaker size {breaker_size}A. Review feeder lengths, connected loads, and breaker size!",
        group="Calculation",
        severity="MEDIUM",
    ),
    "calculation_breaker_lug_quantity_limit": AlertDefinition(
        "Calculations.BreakerLugQuantityLimit",
        "Calculated wire sets ({wire_sets} sets) exceeds recommended maximum for breaker size {breaker_size}A. Review feeder lengths, connected loads, and breaker size!",
        group="Calculation",
        severity="MEDIUM",
    ),
    "design_excessive_conduit_fill": AlertDefinition(
        "Design.ExcessiveConduitFill",
        "User-specified conduit size {conduit_size} exceeds max fill target ({conduit_fill_percentage}% > {max_fill_percentage}%).",
        group="Design",
        severity="MEDIUM",
    ),
    "design_undersized_wire_egc": AlertDefinition(
        "Design.UndersizedWireEGC",
        "User-specified ground size ({ground_size} {wire_material}) does not meet required size per NEC 250.122!",
        group="Design",
        severity="HIGH",
    ),
    "design_undersized_wire_service_ground": AlertDefinition(
        "Design.UndersizedWireServiceGround",
        "User-specified ground size ({ground_size} {wire_material}) does not meet required size per NEC 250.102!",
        group="Design",
        severity="HIGH",
    ),
    "design_excessive_volt_drop": AlertDefinition(
        "Design.ExcessiveVoltDrop",
        "User-specified wire ({wire_sets} set(s) x {wire_size}) fails volt drop check ({vd_percent}%). Review feeder lengths, connected loads and wire sizes!",
        group="Design",
        severity="MEDIUM",
    ),
    "design_insufficient_ampacity": AlertDefinition(
        "Design.InsufficientAmpacity",
        "User-specified wire ({wire_sets} set(s) x {wire_size}) fails ampacity check (Ampacity: {circuit_ampacity}A, Circuit Load: {circuit_load_current}A)",
        group="Design",
        severity="HIGH",
    ),
    "design_insufficient_ampacity_breaker": AlertDefinition(
        "Design.InsufficientAmpacityBreaker",
        "User-specified wire ({wire_sets} set(s) x {wire_size}) fails breaker ampacity check (Ampacity: {circuit_ampacity}A, Breaker: {breaker_rating}A)",
        group="Design",
        severity="HIGH",
    ),
    "design_undersized_ocp": AlertDefinition(
        "Design.UndersizedOCP",
        "Circuit load current ({circuit_load_current}A) exceeds User-specified breaker rating ({breaker_rating}A)",
        group="Design",
        severity="HIGH",
    ),
    "overrides_invalid_isolated_ground": AlertDefinition(
        "Overrides.InvalidIsolatedGround",
        "User-specified isolated ground size '{override_value}' invalid; using equipment ground size instead.",
        group="Overrides",
        severity="NONE",
    ),
    "calculation_wire_sizing_failed": AlertDefinition(
        "Calculations.WireSizingFailed",
        "Unable to automatically calculate wire sizing: {reason}",
        group="Calculation",
        severity="CRITICAL",
    ),
    "calculation_conduit_sizing_failed": AlertDefinition(
        "Calculations.ConduitSizingFailed",
        "Unable to automatically calculate conduit size for fill {fill_ratio}% at max fill {max_fill}%.",
        group="Calculation",
        severity="CRITICAL",
    ),
}
