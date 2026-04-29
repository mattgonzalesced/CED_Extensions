# -*- coding: utf-8 -*-
"""
Shared-parameter setup for the MEPRFP 2.0 ``Element_Linker`` parameter.

The parameter is bound to instance categories that can host placed
equipment children (FamilyInstance and Group). On first use of any
capture tool, ``ensure_element_linker_bound(doc)`` is called; it
swaps the application's shared-parameter file to ours just long
enough to bind the parameter, then restores the prior path.

The parameter GUID is fixed (matches ``_resources/MEPRFP_2_SharedParams.txt``)
so the parameter is recognised across projects.
"""

import os

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    BuiltInParameterGroup,
    Category,
    CategorySet,
    ExternalDefinition,
    InstanceBinding,
)
from System import Guid  # noqa: E402


ELEMENT_LINKER_PARAM_NAME = "Element_Linker"
ELEMENT_LINKER_GUID_STR = "b4e5f6a7-1c2d-3e4f-5a6b-7c8d9e0f1a2b"
ELEMENT_LINKER_GUID = Guid(ELEMENT_LINKER_GUID_STR)
SHARED_PARAM_GROUP_NAME = "MEPRFP_2_Authoring"

_RESOURCES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_resources")
SHARED_PARAM_FILE_PATH = os.path.join(_RESOURCES_DIR, "MEPRFP_2_SharedParams.txt")


# Instance categories that host equipment children. Keep this list
# narrow on purpose — adding categories later is a one-line change;
# binding to every category in the project would surface the parameter
# in unrelated schedules.
_BINDING_CATEGORY_BUILTINS = (
    BuiltInCategory.OST_SpecialityEquipment,
    BuiltInCategory.OST_ElectricalEquipment,
    BuiltInCategory.OST_ElectricalFixtures,
    BuiltInCategory.OST_LightingFixtures,
    BuiltInCategory.OST_LightingDevices,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_MechanicalControlDevices,
    BuiltInCategory.OST_DataDevices,
    BuiltInCategory.OST_PlumbingFixtures,
    BuiltInCategory.OST_FireProtection,
    BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_IOSModelGroups,  # model groups
)


class SharedParamError(Exception):
    pass


def is_element_linker_bound(doc):
    """True if ``Element_Linker`` is already bound on the document."""
    bindings = doc.ParameterBindings
    iterator = bindings.ForwardIterator()
    iterator.Reset()
    while iterator.MoveNext():
        definition = iterator.Key
        if getattr(definition, "Name", None) == ELEMENT_LINKER_PARAM_NAME:
            return True
    return False


def _categories_for_binding(doc):
    cat_set = CategorySet()
    for builtin in _BINDING_CATEGORY_BUILTINS:
        try:
            cat = Category.GetCategory(doc, builtin)
        except Exception:
            cat = None
        if cat is not None and cat.AllowsBoundParameters:
            cat_set.Insert(cat)
    return cat_set


def _open_param_file(app, path):
    prior = app.SharedParametersFilename
    app.SharedParametersFilename = path
    try:
        return prior, app.OpenSharedParameterFile()
    except Exception as exc:
        app.SharedParametersFilename = prior
        raise SharedParamError(
            "Failed to open MEPRFP shared parameter file at {}: {}".format(path, exc)
        )


def _restore_param_file(app, prior):
    app.SharedParametersFilename = prior or ""


def _find_definition(file_obj):
    for group in file_obj.Groups:
        if group.Name != SHARED_PARAM_GROUP_NAME:
            continue
        for definition in group.Definitions:
            if definition.Name == ELEMENT_LINKER_PARAM_NAME:
                return definition
    return None


def ensure_element_linker_bound(doc):
    """Bind ``Element_Linker`` if not already bound. Idempotent.

    Caller is responsible for opening a Revit transaction *before*
    calling this; we mutate ``ParameterBindings``.
    """
    if is_element_linker_bound(doc):
        return False
    if not os.path.isfile(SHARED_PARAM_FILE_PATH):
        raise SharedParamError(
            "MEPRFP shared parameter file is missing: {}".format(SHARED_PARAM_FILE_PATH)
        )
    app = doc.Application
    prior, sp_file = _open_param_file(app, SHARED_PARAM_FILE_PATH)
    try:
        definition = _find_definition(sp_file)
        if definition is None:
            raise SharedParamError(
                "Group/parameter {!r}/{!r} not found in shared parameter file".format(
                    SHARED_PARAM_GROUP_NAME, ELEMENT_LINKER_PARAM_NAME
                )
            )
        category_set = _categories_for_binding(doc)
        if category_set.IsEmpty:
            raise SharedParamError(
                "No bindable categories were found in the active document."
            )
        binding = InstanceBinding(category_set)
        ok = doc.ParameterBindings.Insert(
            definition, binding, BuiltInParameterGroup.PG_DATA
        )
        if not ok:
            raise SharedParamError(
                "Revit refused to bind the {} parameter (Insert returned False).".format(
                    ELEMENT_LINKER_PARAM_NAME
                )
            )
        return True
    finally:
        _restore_param_file(app, prior)
