# -*- coding: utf-8 -*-
try:
    from .PlacementRule import PlacementRule
    from .DynamicParamRule import DynamicParamRule
except Exception:
    try:
        from PlacementRule import PlacementRule
        from DynamicParamRule import DynamicParamRule
    except Exception:
        PlacementRule = None
        DynamicParamRule = None


class LinkedElementDefinition(object):
    def __init__(
        self,
        element_def_id=None,
        category=None,
        family=None,
        type_name=None,
        placement=None,
        static_params=None,
        dynamic_params=None,
        allow_recreate=False,
        is_optional=False,
        is_parent_anchor=False,
    ):
        self._element_def_id = element_def_id
        self._category = category
        self._family = family
        self._type = type_name
        self._placement = placement
        self._static_params = dict(static_params) if static_params is not None else {}
        self._dynamic_params = list(dynamic_params) if dynamic_params is not None else []
        self._allow_recreate = bool(allow_recreate)
        self._is_optional = bool(is_optional)
        self._is_parent_anchor = bool(is_parent_anchor)

    def get_element_def_id(self):
        return self._element_def_id

    def set_element_def_id(self, value):
        self._element_def_id = value

    def get_category(self):
        return self._category

    def set_category(self, value):
        self._category = value

    def get_family(self):
        return self._family

    def set_family(self, value):
        self._family = value

    def get_type(self):
        return self._type

    def set_type(self, value):
        self._type = value

    def get_placement(self):
        return self._placement

    def set_placement(self, value):
        self._placement = value

    def get_static_params(self):
        return self._static_params

    def set_static_params(self, value):
        self._static_params = dict(value) if value is not None else {}

    def get_dynamic_params(self):
        return self._dynamic_params

    def set_dynamic_params(self, value):
        self._dynamic_params = list(value) if value is not None else []

    def get_allow_recreate(self):
        return self._allow_recreate

    def set_allow_recreate(self, value):
        self._allow_recreate = bool(value)

    def get_is_optional(self):
        return self._is_optional

    def set_is_optional(self, value):
        self._is_optional = bool(value)

    def is_parent_anchor(self):
        return self._is_parent_anchor

    def set_is_parent_anchor(self, value):
        self._is_parent_anchor = bool(value)

    # Helpers
    def get_static_param(self, name, default=None):
        return self._static_params.get(name, default)

    def set_static_param(self, name, value):
        self._static_params[name] = value

    def add_dynamic_param(self, rule):
        self._dynamic_params.append(rule)

    def remove_dynamic_param_by_child_param(self, child_param):
        before = len(self._dynamic_params)
        self._dynamic_params = [r for r in self._dynamic_params if not (getattr(r, "get_child_param", lambda: None)() == child_param)]
        return before - len(self._dynamic_params)
