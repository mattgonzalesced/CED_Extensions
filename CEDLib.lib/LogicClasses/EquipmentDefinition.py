# -*- coding: utf-8 -*-
try:
    from .ParentFilter import ParentFilter
    from .EquipmentProperties import EquipmentProperties
    from .LinkedElementSet import LinkedElementSet
except Exception:
    # Fallback if relative imports are not available
    try:
        from ParentFilter import ParentFilter
        from EquipmentProperties import EquipmentProperties
        from LinkedElementSet import LinkedElementSet
    except Exception:
        ParentFilter = None
        EquipmentProperties = None
        LinkedElementSet = None


class EquipmentDefinition(object):
    def __init__(
        self,
        equipment_def_id=None,
        name=None,
        version=None,
        schema_version=None,
        allow_parentless=False,
        allow_unmatched_parents=False,
        prompt_on_parent_mismatch=False,
        parent_filters=None,
        equipment_properties=None,
        linked_sets=None,
    ):
        self._equipment_def_id = equipment_def_id
        self._name = name
        self._version = version
        self._schema_version = schema_version
        self._allow_parentless = bool(allow_parentless)
        self._allow_unmatched_parents = bool(allow_unmatched_parents)
        self._prompt_on_parent_mismatch = bool(prompt_on_parent_mismatch)
        self._parent_filters = parent_filters
        self._equipment_properties = equipment_properties
        self._linked_sets = list(linked_sets) if linked_sets is not None else []

    def get_equipment_def_id(self):
        return self._equipment_def_id

    def set_equipment_def_id(self, value):
        self._equipment_def_id = value

    def get_name(self):
        return self._name

    def set_name(self, value):
        self._name = value

    def get_version(self):
        return self._version

    def set_version(self, value):
        self._version = value

    def get_schema_version(self):
        return self._schema_version

    def set_schema_version(self, value):
        self._schema_version = value

    def get_allow_parentless(self):
        return self._allow_parentless

    def set_allow_parentless(self, value):
        self._allow_parentless = bool(value)

    def get_allow_unmatched_parents(self):
        return self._allow_unmatched_parents

    def set_allow_unmatched_parents(self, value):
        self._allow_unmatched_parents = bool(value)

    def get_prompt_on_parent_mismatch(self):
        return self._prompt_on_parent_mismatch

    def set_prompt_on_parent_mismatch(self, value):
        self._prompt_on_parent_mismatch = bool(value)

    def get_parent_filters(self):
        return self._parent_filters

    def set_parent_filters(self, value):
        self._parent_filters = value

    def get_equipment_properties(self):
        return self._equipment_properties

    def set_equipment_properties(self, value):
        self._equipment_properties = value

    def get_linked_sets(self):
        return self._linked_sets

    def set_linked_sets(self, value):
        self._linked_sets = list(value) if value is not None else []

    # Helpers
    def add_linked_set(self, linked_set):
        """Append a LinkedElementSet."""
        self._linked_sets.append(linked_set)

    def get_linked_set_by_id(self, set_def_id):
        """Return first linked set matching set_def_id or None."""
        for ls in self._linked_sets:
            getter = getattr(ls, "get_set_def_id", None)
            if getter and getter() == set_def_id:
                return ls
        return None

    def remove_linked_set(self, set_def_id):
        """Remove linked sets matching set_def_id; return count removed."""
        before = len(self._linked_sets)
        self._linked_sets = [ls for ls in self._linked_sets if not (getattr(ls, "get_set_def_id", lambda: None)() == set_def_id)]
        return before - len(self._linked_sets)

    def get_all_linked_element_defs(self):
        """Flatten and return all LinkedElementDefinition objects across sets."""
        all_defs = []
        for ls in self._linked_sets:
            getter = getattr(ls, "get_elements", None)
            if getter:
                all_defs.extend(getter() or [])
        return all_defs
