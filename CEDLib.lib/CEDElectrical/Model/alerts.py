# -*- coding: utf-8 -*-
"""Alert definitions and notice collection utilities."""


class AlertDefinition(object):
    """Lightweight alert definition similar to Revit failure definitions."""

    def __init__(self, alert_id, message, group, severity="WARNING", resolutions=None):
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

    def add_message(self, level, message, group="Calculation"):
        self.items.append((None, level.upper(), group, message))

    def add_by_id(self, alert_id, group_override=None, severity_override=None, **fmt):
        from CEDElectrical.refdata.alert_definitions import ALERT_DEFINITIONS

        definition = ALERT_DEFINITIONS.get(alert_id)
        if not definition:
            return
        group = group_override or definition.group
        severity = severity_override or definition.severity
        message = definition.format(**fmt)
        self.items.append((definition, severity.upper(), group, message))

    def has_items(self):
        return bool(self.items)

    def categorized(self):
        """Return items grouped by alert group with severity buckets."""
        ordered = ["Overrides", "Calculation", "Design", "Error", "Other"]
        buckets = {key: {"WARNING": [], "ERROR": []} for key in ordered}
        for definition, severity, group, message in self.items:
            bucket = group if group in buckets else "Other"
            sev = severity.upper() if severity else "WARNING"
            if sev not in buckets[bucket]:
                buckets[bucket][sev] = []
            buckets[bucket][sev].append(message)
        return [(cat, levels) for cat, levels in buckets.items() if levels.get("WARNING") or levels.get("ERROR")]

    def formatted_lines(self, label_map=None):
        if not self.has_items():
            return []
        label_map = label_map or {}
        lines = ["* **{}**".format(self.circuit_name)]
        for category, levels in self.categorized():
            cat_label = label_map.get(category, category)
            cat_msgs = []
            for level in ("ERROR", "WARNING"):
                for msg in levels.get(level, []):
                    cat_msgs.append("    - ({}): {}".format(level.title(), msg))
            if cat_msgs:
                lines.append("  - {}:".format(cat_label))
                lines.extend(cat_msgs)
        return lines


def get_alert_definition(alert_id):
    return ALERT_DEFINITIONS.get(alert_id)
