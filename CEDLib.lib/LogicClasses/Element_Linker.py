# -*- coding: utf-8 -*-
import csv
import json
import os
"""
Element_Linker
--------------
Class-based configuration for CAD → Revit placements.
Designed for IronPython 2.7 / pyRevit / Revit 2024.
"""


class OffsetConfig(object):
    def __init__(self, x_inches=0.0, y_inches=0.0, z_inches=0.0, rotation_deg=0.0):
        self.x_inches = float(x_inches)
        self.y_inches = float(y_inches)
        self.z_inches = float(z_inches)
        self.rotation_deg = float(rotation_deg)


class TagConfig(object):
    """Configuration for a single Generic Annotation (keynote) tag."""

    def __init__(self,
                 category_name,
                 family_name,
                 type_name,
                 parameters=None,
                 offsets=None):
        self.category_name = category_name              # e.g. "Generic Annotations"
        self.family_name = family_name                  # e.g. "Manual Key Note- All Shapes"
        self.type_name = type_name                      # e.g. "Square"
        self.parameters = parameters or {}
        # single OffsetConfig; if you ever want multiple you can change this to a list
        self.offsets = offsets or OffsetConfig(0.0, 0.0, 0.0, 0.0)



class InstanceConfig(object):
    def __init__(self, parameters=None, offsets=None, tags=None):
        """
        parameters: dict of {param_name: value}
        offsets:    OffsetConfig or list[OffsetConfig]
        tags:       list[TagConfig]
        """
        self.parameters = parameters or {}
        if offsets is None:
            self._offsets = []
        elif isinstance(offsets, list):
            self._offsets = offsets
        else:
            self._offsets = [offsets]
        self.tags = tags or []

    def get_offset(self, occurrence_index):
        if not self._offsets:
            return OffsetConfig()
        if occurrence_index < len(self._offsets):
            return self._offsets[occurrence_index]
        # fall back to last entry
        return self._offsets[-1]

    def get_tags(self):
        return self.tags

    def get_parameters(self):
        return self.parameters


class TypeConfig(object):
    def __init__(self, label, category_name, is_group=False, instance_config=None):
        """
        label: "TypeName : FamilyName"  (or "DetailGroup : ModelGroup" for groups)
        category_name: Revit category name string ("Electrical Fixtures", "Model Groups", etc.)
        is_group: True for Model Groups
        instance_config: InstanceConfig
        """
        self.label = label
        self.category_name = category_name
        self.is_group = is_group
        self.instance_config = instance_config or InstanceConfig()


class CadBlockProfile(object):
    """
    All mapping information for a single CAD block name.
    Example: "43 tv", "Stepmill", etc.
    """
    def __init__(self, cad_name):
        self.cad_name = cad_name
        self._types = []  # list[TypeConfig]

    def add_type(self, type_config):
        self._types.append(type_config)

    def add_type_from_parts(self, family_name, type_name, category_name,
                            is_group=False, parameters=None, offsets=None, tags=None):
        """Convenience helper to build and add a TypeConfig from simple parts."""
        label = u"{} : {}".format(family_name, type_name)
        inst_cfg = InstanceConfig(
            parameters=parameters or {},
            offsets=offsets or OffsetConfig(0.0, 0.0, 0.0, 0.0),
            tags=tags or [],
        )
        self._types.append(
            TypeConfig(
                label=label,
                category_name=category_name,
                is_group=is_group,
                instance_config=inst_cfg,
            )
        )

    def get_types(self):
        return list(self._types)

    def get_type_labels(self):
        return [t.label for t in self._types]

    def find_type_by_label(self, label):
        for t in self._types:
            if t.label == label:
                return t
        return None


# -------------------------------------------------------------------------
# Global registry used by the rest of the system.
# This replaces your JSON profile completely.
# -------------------------------------------------------------------------

CAD_BLOCK_PROFILES = {}


def register_profile(profile):
    CAD_BLOCK_PROFILES[profile.cad_name] = profile
    return profile


def register_profiles_from_csv(csv_path, name_columns=None):
    """
    Convenience helper: read a CSV, extract CAD block names, and register
    CadBlockProfile entries for any names not already in CAD_BLOCK_PROFILES.

    csv_path: path to CSV file
    name_columns: optional list of column names (case-insensitive) to try.
                  Defaults to common variants: Name, CAD Name, cad_name, CAD Block, Block.
    Returns: list of cad_name strings that were newly registered.
    """
    cols = name_columns or ["Name", "CAD Name", "cad_name", "CAD Block", "CADBlock", "Block"]
    added = []

    try:
        with open(csv_path, "r") as f:
            reader = csv.DictReader(f)
            for row in reader:
                cad_name = None
                for col in cols:
                    if col.lower() in [k.lower() for k in row.keys()]:
                        # fetch value using original key casing
                        for real_key in row.keys():
                            if real_key.lower() == col.lower():
                                value = row.get(real_key, "")
                                if value:
                                    cad_name = value.strip()
                                break
                    if cad_name:
                        break

                if not cad_name:
                    continue

                if cad_name not in CAD_BLOCK_PROFILES:
                    register_profile(CadBlockProfile(cad_name))
                    added.append(cad_name)
    except Exception:
        # Fail silently but return what we have
        return added

    return added


# -------------------------------------------------------------------------
# Data loading from external YAML (JSON subset).
# -------------------------------------------------------------------------

#DATA_FILE = os.path.join(os.path.dirname(__file__), "element_data.yaml")


#def _offset_from_dict(data):
 #   if data is None:
  #      return OffsetConfig()
   # return OffsetConfig(
    #    data.get("x_inches", 0.0),
     #   data.get("y_inches", 0.0),
      #  data.get("z_inches", 0.0),
       # data.get("rotation_deg", 0.0),
    #)


#def _instance_from_dict(data):
 #   if data is None:
  #      return InstanceConfig()
#
 #   offsets_data = data.get("offsets", []) or []
  #  offsets = []
   # for od in offsets_data:
    #    offsets.append(_offset_from_dict(od))

    #tags = []
    #for tag_data in data.get("tags", []) or []:
     #   tags.append(
      #      TagConfig(
       #         category_name=tag_data.get("category_name"),
        #        family_name=tag_data.get("family_name"),
         #       type_name=tag_data.get("type_name"),
          #      parameters=tag_data.get("parameters") or {},
           #     offsets=_offset_from_dict(tag_data.get("offsets") or {}),
            #)
        #)

    #return InstanceConfig(
     #   parameters=data.get("parameters") or {},
      #  offsets=offsets,
       # tags=tags,
    #)


#def load_profiles_from_yaml(path=None):
#    """Populate CAD_BLOCK_PROFILES from element_data.yaml (stored as JSON-compatible YAML)."""
 #   data_path = path or DATA_FILE
  #  if not os.path.exists(data_path):
   #     raise IOError("Profile data file not found: {}".format(data_path))
#
 #   with open(data_path, "r") as f:
  #      data = json.load(f)
#
 #   CAD_BLOCK_PROFILES.clear()
  #  for prof_data in data.get("profiles", []):
   #     profile = CadBlockProfile(prof_data.get("cad_name"))
    #    for type_data in prof_data.get("types", []) or []:
     #       instance_cfg = _instance_from_dict(type_data.get("instance_config") or {})
      #      type_cfg = TypeConfig(
       #         label=type_data.get("label"),
        #        category_name=type_data.get("category_name"),
         #       is_group=type_data.get("is_group", False),
          #      instance_config=instance_cfg,
           # )
            #profile.add_type(type_cfg)
        #register_profile(profile)


# Load profile data immediately on import
#load_profiles_from_yaml()
