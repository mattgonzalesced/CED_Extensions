# -*- coding: utf-8 -*-
try:
    from .ParentInstance import ParentInstance
    from .ChildLinkInstance import ChildLinkInstance
except Exception:
    try:
        from ParentInstance import ParentInstance
        from ChildLinkInstance import ChildLinkInstance
    except Exception:
        ParentInstance = None
        ChildLinkInstance = None


class ElementLinkInstance(object):
    def __init__(self, equipment_def_id=None, parent=None, child_links=None):
        self._equipment_def_id = equipment_def_id
        self._parent = parent
        self._child_links = list(child_links) if child_links is not None else []

    def get_equipment_def_id(self):
        return self._equipment_def_id

    def set_equipment_def_id(self, value):
        self._equipment_def_id = value

    def get_parent(self):
        return self._parent

    def set_parent(self, value):
        self._parent = value

    def get_child_links(self):
        return self._child_links

    def set_child_links(self, value):
        self._child_links = list(value) if value is not None else []

    # Helpers
    def add_child_link(self, child_link):
        self._child_links.append(child_link)

    def find_child_link(self, element_def_id):
        for cl in self._child_links:
            getter = getattr(cl, "get_element_def_id", None)
            if getter and getter() == element_def_id:
                return cl
        return None

    def remove_child_link(self, element_def_id):
        before = len(self._child_links)
        self._child_links = [cl for cl in self._child_links if not (getattr(cl, "get_element_def_id", lambda: None)() == element_def_id)]
        return before - len(self._child_links)
