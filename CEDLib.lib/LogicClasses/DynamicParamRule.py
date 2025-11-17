# -*- coding: utf-8 -*-


class DynamicParamRule(object):
    def __init__(self, child_param=None, source_type=None, source_value=None):
        self._child_param = child_param
        self._source_type = source_type
        self._source_value = source_value

    def get_child_param(self):
        return self._child_param

    def set_child_param(self, value):
        self._child_param = value

    def get_source_type(self):
        return self._source_type

    def set_source_type(self, value):
        self._source_type = value

    def get_source_value(self):
        return self._source_value

    def set_source_value(self, value):
        self._source_value = value

    # Helpers
    def is_source_type(self, source_type_value):
        """Check if this rule uses the given source_type (string/enum name)."""
        return self._source_type == source_type_value

    def is_literal_value(self):
        """True when source_type suggests a literal value (simple string check)."""
        return str(self._source_type).lower() == "literal"
