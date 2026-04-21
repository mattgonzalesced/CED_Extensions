# -*- coding: utf-8 -*-
"""Shared alert-definition model for CEDElectrical."""


class AlertDefinition(object):
    """Lightweight alert definition similar to Revit failure definitions."""

    def __init__(self, alert_id, message, group, severity="NONE", resolutions=None, persistent=True):
        self._id = alert_id
        self._message = message
        self._group = group
        self._severity = severity
        self._resolutions = resolutions or []
        self._persistent = bool(persistent)

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

    @property
    def persistent(self):
        return self._persistent

    def format(self, **kwargs):
        try:
            return self._message.format(**kwargs)
        except Exception:
            return self._message

