# -*- coding: utf-8 -*-
__title__ = "Parameter Linker"
#from pyrevit import script, revit, DB, forms

class ParameterSet:
    def __init__(self):
        # A list of parameter mappings (pairs of parameters to compare)
        self.parameter_mappings = []  # List of tuples [(param_a, param_b), ...]

    def add_mapping(self, param_a_metadata, param_b_metadata):
        """Adds a parameter mapping to the set."""
        self.parameter_mappings.append((param_a_metadata, param_b_metadata))

    def get_mappings(self):
        """Returns all parameter mappings."""
        return self.parameter_mappings


class ParameterMetadata:
    def __init__(self, name, guid, param_id, storage_type, is_read_only, built_in_param=None):
        self.name = name  # Parameter name
        self.guid = guid  # GUID for shared parameters
        self.param_id = param_id  # Revit Parameter ID
        self.storage_type = storage_type  # StorageType (e.g., Integer, String, etc.)
        self.is_read_only = is_read_only  # Whether the parameter is read-only
        self.built_in_param = built_in_param  # BuiltInParameter enum (if applicable)


    def to_dict(self):
        """Returns a dictionary representation of the metadata."""
        return {
            "name": self.name,
            "guid": str(self.guid) if self.guid else None,
            "param_id": self.param_id,
            "storage_type": self.storage_type,
            "is_read_only": self.is_read_only,
            "built_in_param": self.built_in_param,
        }




# Define metadata for parameters
param_a1 = ParameterMetadata(
    name="FLA Input_CED",
    guid="54564ea7-fc79-44f8-9beb-c9b589901dee",
    param_id=23926625,
    storage_type="Double",
    is_read_only=False,
    built_in_param=False
)

param_a2 = ParameterMetadata(
    name="Voltage_CED",
    guid="04342884-6218-495e-970a-1cdd49f5ddc0",
    param_id=23926634,
    storage_type="Double",
    is_read_only=False,
    built_in_param=False
)

param_a3 = ParameterMetadata(
    name="Phase_CED",
    guid="d4252307-22ba-4917-b756-f79be1334c48",
    param_id=23926632,
    storage_type="Integer",
    is_read_only=True,
    built_in_param=False
)

param_b1 = ParameterMetadata(
    name="CED-E-FLA",
    guid=None,
    param_id=2001,
    storage_type="Double",
    is_read_only=False,
    built_in_param=False
)

param_b2 = ParameterMetadata(
    name="VOLTAGE",
    guid=None,
    param_id=2002,
    storage_type="String",
    is_read_only=False,
    built_in_param=False
)

param_b3 = ParameterMetadata(
    name="PHASE",
    guid=None,
    param_id=2003,
    storage_type="Integer",
    is_read_only=False,
    built_in_param=False
)

# Create a ParameterSet and add mappings
parameter_set = ParameterSet()
parameter_set.add_mapping(param_a1, param_b1)
parameter_set.add_mapping(param_a2, param_b2)
parameter_set.add_mapping(param_a3, param_b3)

# Retrieve and display mappings
for param_a, param_b in parameter_set.get_mappings():
    print("Mapping:")
    print("  Element A - Parameter:", param_a.to_dict())
    print("  Element B - Parameter:", param_b.to_dict())

for params_a, params_b in parameter_set.get_mappings():
    params_a.to_dict()