# -*- coding: utf-8 -*-
try:
    from .WarningItem import WarningItem
except Exception:
    try:
        from WarningItem import WarningItem
    except Exception:
        WarningItem = None


class ChildLinkInstance(object):
    def __init__(
        self,
        element_id=None,
        element_def_id=None,
        set_def_id=None,
        last_known_location=None,
        last_known_type=None,
        warnings=None,
    ):
        self._element_id = element_id
        self._element_def_id = element_def_id
        self._set_def_id = set_def_id
        self._last_known_location = tuple(last_known_location) if last_known_location is not None else None
        self._last_known_type = last_known_type
        self._warnings = list(warnings) if warnings is not None else []

    def get_element_id(self):
        return self._element_id

    def set_element_id(self, value):
        self._element_id = value

    def get_element_def_id(self):
        return self._element_def_id

    def set_element_def_id(self, value):
        self._element_def_id = value

    def get_set_def_id(self):
        return self._set_def_id

    def set_set_def_id(self, value):
        self._set_def_id = value

    def get_last_known_location(self):
        return self._last_known_location

    def set_last_known_location(self, value):
        self._last_known_location = tuple(value) if value is not None else None

    def get_last_known_type(self):
        return self._last_known_type

    def set_last_known_type(self, value):
        self._last_known_type = value

    def get_warnings(self):
        return self._warnings

    def set_warnings(self, value):
        self._warnings = list(value) if value is not None else []

    # Helpers
    def add_warning(self, warning_item):
        """Append a WarningItem-like object."""
        self._warnings.append(warning_item)

    def clear_warnings(self):
        self._warnings = []
