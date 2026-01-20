# -*- coding: utf-8 -*-
try:
    from .LinkedElementDefinition import LinkedElementDefinition
except Exception:
    try:
        from LinkedElementDefinition import LinkedElementDefinition
    except Exception:
        LinkedElementDefinition = None


class LinkedElementSet(object):
    def __init__(self, set_def_id=None, name=None, elements=None):
        self._set_def_id = set_def_id
        self._name = name
        self._elements = list(elements) if elements is not None else []

    def get_set_def_id(self):
        return self._set_def_id

    def set_set_def_id(self, value):
        self._set_def_id = value

    def get_name(self):
        return self._name

    def set_name(self, value):
        self._name = value

    def get_elements(self):
        return self._elements

    def set_elements(self, value):
        self._elements = list(value) if value is not None else []

    # Helpers
    def add_element(self, element_def):
        self._elements.append(element_def)

    def find_element(self, element_def_id):
        for el in self._elements:
            getter = getattr(el, "get_element_def_id", None)
            if getter and getter() == element_def_id:
                return el
        return None

    def remove_element(self, element_def_id):
        before = len(self._elements)
        self._elements = [el for el in self._elements if not (getattr(el, "get_element_def_id", lambda: None)() == element_def_id)]
        return before - len(self._elements)
