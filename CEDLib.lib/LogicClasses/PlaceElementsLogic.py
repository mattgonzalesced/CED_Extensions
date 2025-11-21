# -*- coding: utf-8 -*-
"""
PlaceElementsLogic
------------------
Compatibility shim that re-exports the core pieces used by the Place Elements tool.

The actual implementation now lives in smaller modules under LogicClasses:

* csv_helpers.py       -> feet_inch_to_inches, read_xyz_csv
* profile_repository.py -> ProfileRepository
* placement_engine.py   -> PlaceElementsEngine
* tag_utils.py          -> tag_key_from_dict
"""

from LogicClasses.csv_helpers import feet_inch_to_inches, read_xyz_csv
from LogicClasses.profile_repository import ProfileRepository
from LogicClasses.placement_engine import PlaceElementsEngine
from LogicClasses.tag_utils import tag_key_from_dict

__all__ = [
    "feet_inch_to_inches",
    "read_xyz_csv",
    "ProfileRepository",
    "PlaceElementsEngine",
    "tag_key_from_dict",
]
