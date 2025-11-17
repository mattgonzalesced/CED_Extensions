# -*- coding: utf-8 -*-


class ParentInstance(object):
    def __init__(
        self,
        element_id=None,
        category=None,
        family=None,
        type_name=None,
        location=None,
        is_view_specific=False,
        view_id=None,
    ):
        self._element_id = element_id
        self._category = category
        self._family = family
        self._type = type_name
        self._location = tuple(location) if location is not None else None
        self._is_view_specific = bool(is_view_specific)
        self._view_id = view_id

    def get_element_id(self):
        return self._element_id

    def set_element_id(self, value):
        self._element_id = value

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

    def get_location(self):
        return self._location

    def set_location(self, value):
        self._location = tuple(value) if value is not None else None

    def get_is_view_specific(self):
        return self._is_view_specific

    def set_is_view_specific(self, value):
        self._is_view_specific = bool(value)

    def get_view_id(self):
        return self._view_id

    def set_view_id(self, value):
        self._view_id = value

    # Helpers
    def update_location(self, new_location, view_id=None):
        """Set a new location (tuple) and optionally view_id."""
        self._location = tuple(new_location) if new_location is not None else None
        if view_id is not None:
            self._view_id = view_id
