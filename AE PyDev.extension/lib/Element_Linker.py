# -*- coding: utf-8 -*-
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
    def __init__(self, label, offset=None):
        """
        label: "TagTypeName : TagFamilyName"
        offset: OffsetConfig
        """
        self.label = label
        self.offset = offset or OffsetConfig()


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


# -------------------------------------------------------------------------
# Example definitions
# (You will expand these to match your real mappings)
# -------------------------------------------------------------------------

# 1) Exit Sign ------------------------------------------------------------

profile_exit_sign = register_profile(CadBlockProfile("Exit Sign"))

profile_exit_sign.add_type(
    TypeConfig(
        label="LF-U_Exit Sign_CED : E1",  # Family : Type
        category_name="Lighting Fixtures",
        is_group=False,
        instance_config=InstanceConfig(
            parameters={
                "dev-Group ID": "E1",
                "FLA Input_CED": "0",
                "Apparent Load Input_CED": "180",
                "Load Classification_CED": "L",
                "Voltage_CED": "120",
                "Number of Poles_CED": "1",
                "CKT_Panel_CEDT": "L3",
                "CKT_Circuit Number_CEDT": "EMERGENCY",
                "CKT_Rating_CED": "20",
                "CKT_Load Name_CEDT": "EMERGENCY - (Space)",
                "CKT_Schedule Notes_CEDT": "",
            },
            offsets=OffsetConfig(0.0, 0.0, 0.0, 0.0),
            tags=[],
        ),
    )
)

# 2) Exit Directional Sign ------------------------------------------------

profile_exit_directional_sign = register_profile(
    CadBlockProfile("Exit Directional Sign")
)

profile_exit_directional_sign.add_type(
    TypeConfig(
        label="LF-U_Exit Sign_CED : E2",  # Family : Type
        category_name="Lighting Fixtures",
        is_group=False,
        instance_config=InstanceConfig(
            parameters={
                "dev-Group ID": "E2",
                "FLA Input_CED": "0",
                "Apparent Load Input_CED": "180",
                "Load Classification_CED": "L",
                "Voltage_CED": "120",
                "Number of Poles_CED": "1",
                "CKT_Panel_CEDT": "L3",
                "CKT_Circuit Number_CEDT": "EMERGENCY",
                "CKT_Rating_CED": "20",
                "CKT_Load Name_CEDT": "EMERGENCY - (Space)",
                "CKT_Schedule Notes_CEDT": "",
            },
            offsets=OffsetConfig(0.0, 0.0, 0.0, 0.0),
            tags=[],
        ),
    )
)

# 3) PF_Light Recess ------------------------------------------------------

profile_pf_light_recess = register_profile(CadBlockProfile("PF_Light Recess"))

profile_pf_light_recess.add_type(
    TypeConfig(
        label="LF-U_Round Fixture_CED : F1A",  # Family : Type
        category_name="Lighting Fixtures",
        is_group=False,
        instance_config=InstanceConfig(
            parameters={
                "dev-Group ID": "F1A and F1B",
                "FLA Input_CED": "0",
                "Apparent Load Input_CED": "180",
                "Load Classification_CED": "L",
                "Voltage_CED": "120",
                "Number of Poles_CED": "1",
                "CKT_Panel_CEDT": "L3",
                "CKT_Circuit Number_CEDT": "STANDARD",
                "CKT_Rating_CED": "20",
                "CKT_Load Name_CEDT": "(Space)",
                "CKT_Schedule Notes_CEDT": "",
            },
            offsets=OffsetConfig(0.0, 0.0, 0.0, 0.0),
            tags=[],
        ),
    )
)

# 4) PF_Light Exterior ----------------------------------------------------

profile_pf_light_exterior = register_profile(CadBlockProfile("PF_Light Exterior"))

profile_pf_light_exterior.add_type(
    TypeConfig(
        label="LF-U_Bug Eye Fixture_CED : F13",  # Family : Type
        category_name="Lighting Fixtures",
        is_group=False,
        instance_config=InstanceConfig(
            parameters={
                "dev-Group ID": "F13",
                "FLA Input_CED": "0",
                "Apparent Load Input_CED": "180",
                "Load Classification_CED": "L",
                "Voltage_CED": "120",
                "Number of Poles_CED": "1",
                "CKT_Panel_CEDT": "L3",
                "CKT_Circuit Number_CEDT": "STANDARD",
                "CKT_Rating_CED": "20",
                "CKT_Load Name_CEDT": "(Space)",
                "CKT_Schedule Notes_CEDT": "",
            },
            offsets=OffsetConfig(0.0, 0.0, 0.0, 0.0),
            tags=[],
        ),
    )
)

# 5) PF_Light Surface Mtd -------------------------------------------------

profile_pf_light_surface_mtd = register_profile(
    CadBlockProfile("PF_Light Surface Mtd")
)

profile_pf_light_surface_mtd.add_type(
    TypeConfig(
        label="LF-U_Rectangular Fixture_CED : F9",  # Family : Type
        category_name="Lighting Fixtures",
        is_group=False,
        instance_config=InstanceConfig(
            parameters={
                "dev-Group ID": "F9",
                "FLA Input_CED": "0",
                "Apparent Load Input_CED": "180",
                "Load Classification_CED": "L",
                "Voltage_CED": "120",
                "Number of Poles_CED": "1",
                "CKT_Panel_CEDT": "L3",
                "CKT_Circuit Number_CEDT": "STANDARD",
                "CKT_Rating_CED": "20",
                "CKT_Load Name_CEDT": "(Space)",
                "CKT_Schedule Notes_CEDT": "",
            },
            offsets=OffsetConfig(0.0, 0.0, 0.0, 0.0),
            tags=[],
        ),
    )
)

# 6) PF_Light Sconce ------------------------------------------------------

profile_pf_light_sconce = register_profile(CadBlockProfile("PF_Light Sconce"))

profile_pf_light_sconce.add_type(
    TypeConfig(
        label="LF-U_Wall Sconce Fixture_CED : F12",  # Family : Type
        category_name="Lighting Fixtures",
        is_group=False,
        instance_config=InstanceConfig(
            parameters={
                "dev-Group ID": "F12",
                "FLA Input_CED": "0",
                "Apparent Load Input_CED": "180",
                "Load Classification_CED": "L",
                "Voltage_CED": "120",
                "Number of Poles_CED": "1",
                "CKT_Panel_CEDT": "L3",
                "CKT_Circuit Number_CEDT": "STANDARD",
                "CKT_Rating_CED": "20",
                "CKT_Load Name_CEDT": "(Space)",
                "CKT_Schedule Notes_CEDT": "",
            },
            offsets=OffsetConfig(0.0, 0.0, 0.0, 0.0),
            tags=[],
        ),
    )
)

# 7) PF_Light Fluorescent -------------------------------------------------

profile_pf_light_fluorescent = register_profile(
    CadBlockProfile("PF_Light Fluorescent")
)

profile_pf_light_fluorescent.add_type(
    TypeConfig(
        label="LF-U_Rectangular Fixture_CED : F10",  # Family : Type
        category_name="Lighting Fixtures",
        is_group=False,
        instance_config=InstanceConfig(
            parameters={
                "dev-Group ID": "F10",
                "FLA Input_CED": "0",
                "Apparent Load Input_CED": "180",
                "Load Classification_CED": "L",
                "Voltage_CED": "120",
                "Number of Poles_CED": "1",
                "CKT_Panel_CEDT": "L3",
                "CKT_Circuit Number_CEDT": "STANDARD",
                "CKT_Rating_CED": "20",
                "CKT_Load Name_CEDT": "(Space)",
                "CKT_Schedule Notes_CEDT": "",
            },
            offsets=OffsetConfig(0.0, 0.0, 0.0, 0.0),
            tags=[],
        ),
    )
)

# 8) PF_Light Podium Desk -------------------------------------------------

profile_pf_light_podium_desk = register_profile(
    CadBlockProfile("PF_Light Podium Desk")
)

profile_pf_light_podium_desk.add_type(
    TypeConfig(
        label="LF-U_Rectangular Fixture_CED : F6",  # Family : Type
        category_name="Lighting Fixtures",
        is_group=False,
        instance_config=InstanceConfig(
            parameters={
                "dev-Group ID": "F6",
                "FLA Input_CED": "0",
                "Apparent Load Input_CED": "180",
                "Load Classification_CED": "L",
                "Voltage_CED": "120",
                "Number of Poles_CED": "1",
                "CKT_Panel_CEDT": "L3",
                "CKT_Circuit Number_CEDT": "STANDARD",
                "CKT_Rating_CED": "20",
                "CKT_Load Name_CEDT": "(Space)",
                "CKT_Schedule Notes_CEDT": "",
            },
            offsets=OffsetConfig(0.0, 0.0, 0.0, 0.0),
            tags=[],
        ),
    )
)

# 9) PF_Light Linear Bay --------------------------------------------------

profile_pf_light_linear_bay = register_profile(
    CadBlockProfile("PF_Light Linear Bay")
)

profile_pf_light_linear_bay.add_type(
    TypeConfig(
        label="LF-U_Universal Light Fixture_CED : F2",  # Family : Type
        category_name="Lighting Fixtures",
        is_group=False,
        instance_config=InstanceConfig(
            parameters={
                "dev-Group ID": "F2",
                "FLA Input_CED": "0",
                "Apparent Load Input_CED": "180",
                "Load Classification_CED": "L",
                "Voltage_CED": "120",
                "Number of Poles_CED": "1",
                "CKT_Panel_CEDT": "L3",
                "CKT_Circuit Number_CEDT": "STANDARD",
                "CKT_Rating_CED": "20",
                "CKT_Load Name_CEDT": "(Space)",
                "CKT_Schedule Notes_CEDT": "",
            },
            offsets=OffsetConfig(0.0, 0.0, 0.0, 0.0),
            tags=[],
        ),
    )
)
