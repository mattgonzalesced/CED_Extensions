# -*- coding: utf-8 -*-


class ParentFilter(object):
    def __init__(
        self,
        category=None,
        family_name_pattern=None,
        type_name_pattern=None,
        parameter_filters=None,
    ):
        self._category = category
        self._family_name_pattern = family_name_pattern
        self._type_name_pattern = type_name_pattern
        self._parameter_filters = dict(parameter_filters) if parameter_filters is not None else {}

    def get_category(self):
        return self._category

    def set_category(self, value):
        self._category = value

    def get_family_name_pattern(self):
        return self._family_name_pattern

    def set_family_name_pattern(self, value):
        self._family_name_pattern = value

    def get_type_name_pattern(self):
        return self._type_name_pattern

    def set_type_name_pattern(self, value):
        self._type_name_pattern = value

    def get_parameter_filters(self):
        return self._parameter_filters

    def set_parameter_filters(self, value):
        self._parameter_filters = dict(value) if value is not None else {}

    # Helpers
    def matches_basic(self, category, family_name, type_name):
        """Simple match helper using case-insensitive equality if a pattern is provided."""
        if self._category and (category or "").lower() != (self._category or "").lower():
            return False
        if self._family_name_pattern and (family_name or "").lower() != (self._family_name_pattern or "").lower():
            return False
        if self._type_name_pattern and (type_name or "").lower() != (self._type_name_pattern or "").lower():
            return False
        return True

    def matches_parameters(self, param_dict):
        """Return True if all parameter_filters match exactly in param_dict."""
        for key, val in (self._parameter_filters or {}).items():
            if param_dict.get(key) != val:
                return False
        return True
