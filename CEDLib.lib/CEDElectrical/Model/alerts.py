# -*- coding: utf-8 -*-
"""Alert definitions and notice collection utilities."""


class AlertDefinition(object):
    """Lightweight alert definition similar to Revit failure definitions."""

    def __init__(self, alert_id, message, group, severity="NONE", resolutions=None):
        self._id = alert_id
        self._message = message
        self._group = group
        self._severity = severity
        self._resolutions = resolutions or []

    # Accessor-style methods mirroring Revit's API naming
    def GetApplicableResolutionTypes(self):
        return list(self._resolutions)

    def GetDefaultResolutionType(self):
        return self._resolutions[0] if self._resolutions else None

    def GetDescriptionText(self):
        return self._message

    def GetId(self):
        return self._id

    def GetSeverity(self):
        return self._severity

    def HasResolutions(self):
        return bool(self._resolutions)

    @property
    def group(self):
        return self._group

    @property
    def message(self):
        return self._message

    @property
    def severity(self):
        return self._severity

    def format(self, **kwargs):
        try:
            return self._message.format(**kwargs)
        except Exception:
            return self._message


class NoticeCollector(object):
    """Collects alerts for a branch so they can be summarized later."""

    def __init__(self, circuit_name):
        self.circuit_name = circuit_name
        self.items = []  # (AlertDefinition, severity, group, message)

    def add_message(self, severity, message, group="Calculation"):
        sev = (severity or "NONE").upper()
        self.items.append((None, sev, group, message))

    def add_alert(self, alert_spec, group_override=None, severity_override=None):
        if not alert_spec:
            return

        definition = alert_spec.get("definition") if isinstance(alert_spec, dict) else None
        if not definition:
            return

        data = alert_spec.get("data", {}) if isinstance(alert_spec, dict) else {}
        group = group_override or definition.group
        severity = (severity_override or definition.severity or "NONE").upper()
        message = definition.format(**data)
        self.items.append((definition, severity, group, message))

    def add_by_id(self, alert_id, group_override=None, severity_override=None, **fmt):
        definition = get_alert_definition(alert_id)
        if not definition:
            return
        self.add_alert({"definition": definition, "data": fmt}, group_override, severity_override)

    def has_items(self):
        return bool(self.items)

    def formatted_lines(self, label_map=None, severity_colors=None):
        if not self.has_items():
            return []
        label_map = label_map or {}
        severity_colors = severity_colors or {}
        lines = ["**{}**".format(self.circuit_name)]

        for _, severity, group, message in self.items:
            label = label_map.get(group, group)
            sev_key = severity.upper() if severity else "NONE"
            color = severity_colors.get(sev_key)
            rendered = message
            if color:
                rendered = "<span style=\"color:{}\">{}</span>".format(color, message)
            lines.append("  - (**{}**) {}".format(label, rendered))

        return lines


class Alerts(object):
    """Static constructors for typed alert specs."""

    @staticmethod
    def InvalidCircuitProperty(property, override_value, default_value):
        return {
            "definition": get_alert_definition("overrides_invalid_circuit_property"),
            "data": {
                "property": property,
                "override_value": override_value,
                "default_value": default_value,
            },
        }

    @staticmethod
    def InvalidEquipmentGround(override_value):
        return {
            "definition": get_alert_definition("overrides_invalid_equipment_ground"),
            "data": {"override_value": override_value},
        }

    @staticmethod
    def InvalidServiceGround(override_value):
        return {
            "definition": get_alert_definition("overrides_invalid_service_ground"),
            "data": {"override_value": override_value},
        }

    @staticmethod
    def InvalidHotWire(override_value):
        return {
            "definition": get_alert_definition("overrides_invalid_hot_wire"),
            "data": {"override_value": override_value},
        }

    @staticmethod
    def InvalidConduit(override_value):
        return {
            "definition": get_alert_definition("overrides_invalid_conduit"),
            "data": {"override_value": override_value},
        }

    @staticmethod
    def NonStandardOCPRating(breaker_size, next_size):
        return {
            "definition": get_alert_definition("design_non_standard_ocp_rating"),
            "data": {"breaker_size": breaker_size, "next_size": next_size},
        }

    @staticmethod
    def BreakerLugSizeLimitOverride(hot_size, breaker_size, max_lug_size):
        return {
            "definition": get_alert_definition("design_breaker_lug_size_limit_override"),
            "data": {
                "hot_size": hot_size,
                "breaker_size": breaker_size,
                "max_lug_size": max_lug_size,
            },
        }

    @staticmethod
    def BreakerLugQuantityLimitOverride(wire_sets, breaker_size, max_lug_qty):
        return {
            "definition": get_alert_definition("design_breaker_lug_quantity_limit_override"),
            "data": {
                "wire_sets": wire_sets,
                "breaker_size": breaker_size,
                "max_lug_qty": max_lug_qty,
            },
        }

    @staticmethod
    def BreakerLugSizeLimitCalc(hot_size, breaker_size):
        return {
            "definition": get_alert_definition("calculation_breaker_lug_size_limit"),
            "data": {"hot_size": hot_size, "breaker_size": breaker_size},
        }

    @staticmethod
    def BreakerLugQuantityLimitCalc(wire_sets, breaker_size):
        return {
            "definition": get_alert_definition("calculation_breaker_lug_quantity_limit"),
            "data": {"wire_sets": wire_sets, "breaker_size": breaker_size},
        }

    @staticmethod
    def ExcessiveConduitFill(conduit_size, conduit_fill_percentage, max_fill_percentage):
        return {
            "definition": get_alert_definition("design_excessive_conduit_fill"),
            "data": {
                "conduit_size": conduit_size,
                "conduit_fill_percentage": conduit_fill_percentage,
                "max_fill_percentage": max_fill_percentage,
            },
        }

    @staticmethod
    def UndersizedWireEGC(ground_size, wire_material):
        return {
            "definition": get_alert_definition("design_undersized_wire_egc"),
            "data": {"ground_size": ground_size, "wire_material": wire_material},
        }

    @staticmethod
    def UndersizedWireServiceGround(ground_size, wire_material):
        return {
            "definition": get_alert_definition("design_undersized_wire_service_ground"),
            "data": {"ground_size": ground_size, "wire_material": wire_material},
        }

    @staticmethod
    def ExcessiveVoltDrop(wire_sets, wire_size, vd_percent):
        return {
            "definition": get_alert_definition("design_excessive_volt_drop"),
            "data": {
                "wire_sets": wire_sets,
                "wire_size": wire_size,
                "vd_percent": vd_percent,
            },
        }

    @staticmethod
    def InsufficientAmpacity(wire_sets, wire_size, circuit_ampacity, circuit_load_current):
        return {
            "definition": get_alert_definition("design_insufficient_ampacity"),
            "data": {
                "wire_sets": wire_sets,
                "wire_size": wire_size,
                "circuit_ampacity": circuit_ampacity,
                "circuit_load_current": circuit_load_current,
            },
        }

    @staticmethod
    def UndersizedOCP(circuit_load_current, breaker_rating):
        return {
            "definition": get_alert_definition("design_undersized_ocp"),
            "data": {
                "circuit_load_current": circuit_load_current,
                "breaker_rating": int(round(breaker_rating)),
            },
        }

    @staticmethod
    def WireSizingFailed(reason):
        return {
            "definition": get_alert_definition("calculation_wire_sizing_failed"),
            "data": {"reason": reason},
        }

    @staticmethod
    def ConduitSizingFailed(fill_ratio, max_fill):
        return {
            "definition": get_alert_definition("calculation_conduit_sizing_failed"),
            "data": {"fill_ratio": fill_ratio, "max_fill": max_fill},
        }

def get_alert_definition(alert_id):
    try:
        from CEDElectrical.refdata.alert_definitions import ALERT_DEFINITIONS
    except Exception:
        return None

    definition = ALERT_DEFINITIONS.get(alert_id)
    if definition:
        return definition

    # fall back to searching by definition id
    for defn in ALERT_DEFINITIONS.values():
        try:
            if defn.GetId() == alert_id:
                return defn
        except Exception:
            continue
    return None
