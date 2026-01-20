# -*- coding: utf-8 -*-


class EquipmentProperties(object):
    def __init__(
        self,
        manufacturer=None,
        model_number=None,
        description=None,
        dimensions=None,
        electrical=None,
        mechanical=None,
        weight=None,
        other=None,
    ):
        self._manufacturer = manufacturer
        self._model_number = model_number
        self._description = description
        self._dimensions = dict(dimensions) if dimensions is not None else {}
        self._electrical = dict(electrical) if electrical is not None else {}
        self._mechanical = dict(mechanical) if mechanical is not None else {}
        self._weight = weight
        self._other = dict(other) if other is not None else {}

    def get_manufacturer(self):
        return self._manufacturer

    def set_manufacturer(self, value):
        self._manufacturer = value

    def get_model_number(self):
        return self._model_number

    def set_model_number(self, value):
        self._model_number = value

    def get_description(self):
        return self._description

    def set_description(self, value):
        self._description = value

    def get_dimensions(self):
        return self._dimensions

    def set_dimensions(self, value):
        self._dimensions = dict(value) if value is not None else {}

    def get_electrical(self):
        return self._electrical

    def set_electrical(self, value):
        self._electrical = dict(value) if value is not None else {}

    def get_mechanical(self):
        return self._mechanical

    def set_mechanical(self, value):
        self._mechanical = dict(value) if value is not None else {}

    def get_weight(self):
        return self._weight

    def set_weight(self, value):
        self._weight = value

    def get_other(self):
        return self._other

    def set_other(self, value):
        self._other = dict(value) if value is not None else {}

    # Helpers
    def merge_from(self, other):
        """Shallow-merge fields from another EquipmentProperties-like object."""
        if other is None:
            return
        for attr in ["manufacturer", "model_number", "description", "weight"]:
            getter = getattr(other, "get_{0}".format(attr), None)
            if getter:
                val = getter()
                if val is not None:
                    setattr(self, "_{0}".format(attr), val)
        # Merge dict-like fields
        for attr in ["dimensions", "electrical", "mechanical", "other"]:
            getter = getattr(other, "get_{0}".format(attr), None)
            if getter:
                incoming = getter() or {}
                current = getattr(self, "_{0}".format(attr))
                current.update(incoming)

    def to_dict(self):
        """Return a basic dict representation."""
        return {
            "manufacturer": self._manufacturer,
            "model_number": self._model_number,
            "description": self._description,
            "dimensions": dict(self._dimensions),
            "electrical": dict(self._electrical),
            "mechanical": dict(self._mechanical),
            "weight": self._weight,
            "other": dict(self._other),
        }
