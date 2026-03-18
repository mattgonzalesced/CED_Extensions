"""Export family configurations to yaml file.

Family configuration file is a yaml file,
providing info about the parameters and types defined in the family.
The shared parameters are exported to a txt file.
In the yaml file, the shared parameters are distinguished
by the presence of their GUID.

The structure of this config file is as shown below:

parameters:
    <parameter-name>:
        type: <Autodesk.Revit.DB.ParameterType> or
        <Autodesk.Revit.DB.ParameterTypeId Members> (2022+)
        group: <Autodesk.Revit.DB.BuiltInParameterGroup> or
        <Autodesk.Revit.DB.GroupTypeId Members> (2022+)
        instance: <true|false>
        reporting: <true|false>
        formula: <str>
        default: <str>
types:
    <type-name>:
        <parameter-name>: <value>
        <parameter-name>: <value>
        ...


Example:

parameters:
    Shelf Height (Upper):
        type: Length
        group: PG_GEOMETRY or Geometry (2022+)
        instance: false
types:
    24D"x36H":
        Shelf Height (Upper): 3'-0"

Note: If a parameter is in the revit file and the yaml file,
but shared in one and family in the other, after import,
the parameter won't change. So if it was shared in the revit file,
but family in the yaml file, it will remain shared.
"""
# pylint: disable=import-error,invalid-name,broad-except
# FIXME export parameter ordering

import os

from pyrevit import HOST_APP
from pyrevit import revit, DB
from pyrevit import forms
from pyrevit import coreutils
from pyrevit import script
from pyrevit.coreutils import yaml

from Autodesk.Revit import Exceptions


logger = script.get_logger()
output = script.get_output()


# yaml sections and keys ------------------------------------------------------
PARAM_SECTION_NAME = 'parameters'
PARAM_SECTION_TYPE = 'type'
PARAM_SECTION_CAT = 'category'
PARAM_SECTION_GROUP = 'group'
PARAM_SECTION_INST = 'instance'
PARAM_SECTION_REPORT = 'reporting'
PARAM_SECTION_FORMULA = 'formula'
PARAM_SECTION_DEFAULT = 'default'
PARAM_SECTION_GUID = 'GUID'  # To store unique if of shared parameters

TYPES_SECTION_NAME = 'types'

SHAREDPARAM_DEF = 'xref_sharedparams'
# -----------------------------------------------------------------------------

FAMILY_SYMBOL_FORMAT = '{} : {}'
ELECTRICAL_LOAD_CLASS_FORMAT = 'ELECTRICAL_LOAD_CLASS : {}'
MAP_FILE_NAME = 'map.yaml'
FAMILY_CONFIGS_ROOT = r'C:\ACC\ACCDocs\CoolSys\CED Content Collection\Project Files\03 Automations\Family Configs'
EXPORT_SOURCE_DIR = os.path.join(FAMILY_CONFIGS_ROOT, 'Source')

ACTIVE_DOC = revit.doc
SOURCE_DOC = revit.doc


def _get_active_doc():
    return ACTIVE_DOC

def _get_source_doc():
    return SOURCE_DOC


def _clean_filename_token(value, fallback):
    token = (value or '').strip()
    if not token:
        token = fallback
    for bad_char in '<>:"/\\|?*':
        token = token.replace(bad_char, '-')
    token = token.rstrip('. ').strip()
    return token or fallback


def _get_project_info_doc():
    source_doc = _get_source_doc()
    if source_doc and not source_doc.IsFamilyDocument:
        return source_doc
    active_doc = _get_active_doc()
    if active_doc and not active_doc.IsFamilyDocument:
        return active_doc
    return source_doc or active_doc


def _read_projectinfo_value(project_doc, attr_name, bip_name):
    if not project_doc:
        return ''
    try:
        pinfo = project_doc.ProjectInformation
    except Exception:
        pinfo = None
    if not pinfo:
        return ''

    try:
        attr_value = getattr(pinfo, attr_name, '')
        if attr_value:
            return str(attr_value).strip()
    except Exception:
        pass

    try:
        bip = getattr(DB.BuiltInParameter, bip_name)
        param = pinfo.get_Parameter(bip)
        if param:
            return (param.AsString() or '').strip()
    except Exception:
        pass

    return ''


def _get_default_export_filename():
    family_name = _clean_filename_token(_get_family_name(), 'Family')
    project_doc = _get_project_info_doc()
    client_name = _clean_filename_token(
        _read_projectinfo_value(project_doc, 'ClientName', 'PROJECT_CLIENT_NAME'),
        'UnknownClient'
    )
    project_number = _clean_filename_token(
        _read_projectinfo_value(project_doc, 'Number', 'PROJECT_NUMBER'),
        'UnknownProjectNumber'
    )
    return '{}_{}_{}.yaml'.format(family_name, client_name, project_number)


def _doc_signature(doc):
    try:
        return (doc.GetHashCode(), doc.Title)
    except Exception:
        return (id(doc), getattr(doc, 'Title', repr(doc)))


def _get_open_doc_signatures():
    sigs = set()
    for open_doc in HOST_APP.app.Documents:
        try:
            sigs.add(_doc_signature(open_doc))
        except Exception:
            continue
    return sigs


def _get_open_family_doc_for(target_family):
    if not target_family:
        return None

    try:
        target_id = target_family.Id.IntegerValue
    except Exception:
        target_id = None

    for open_doc in HOST_APP.app.Documents:
        try:
            if not open_doc.IsFamilyDocument:
                continue
            owner_family = open_doc.FamilyManager.OwnerFamily
            if owner_family and owner_family.Id.IntegerValue == target_id:
                return open_doc
        except Exception:
            continue
    return None


def _pick_loaded_family(project_doc):
    families = list(DB.FilteredElementCollector(project_doc).OfClass(DB.Family))
    if not families:
        forms.alert('No loaded families were found in this project.', warn_icon=True)
        return None

    names = [f.Name or '<Unnamed Family>' for f in families]
    name_counts = {}
    for name in names:
        name_counts[name] = name_counts.get(name, 0) + 1

    label_to_family = {}
    labels = []
    for family in families:
        family_name = family.Name or '<Unnamed Family>'
        if name_counts.get(family_name, 0) > 1:
            label = '{} (Id {})'.format(family_name, family.Id.IntegerValue)
        else:
            label = family_name
        labels.append(label)
        label_to_family[label] = family

    labels = sorted(labels, key=lambda x: x.lower())
    selected_label = forms.SelectFromList.show(
        labels,
        title='Select Family to Export',
        button_name='Select Family',
        multiselect=False
    )
    if not selected_label:
        return None

    return label_to_family.get(selected_label)


def _resolve_family_doc_for_export():
    project_or_family_doc = revit.doc
    if project_or_family_doc.IsFamilyDocument:
        return project_or_family_doc, False

    target_family = _pick_loaded_family(project_or_family_doc)
    if not target_family:
        return None, False

    open_family_doc = _get_open_family_doc_for(target_family)
    if open_family_doc:
        return open_family_doc, False

    existing_doc_sigs = _get_open_doc_signatures()
    family_doc = project_or_family_doc.EditFamily(target_family)
    should_close = _doc_signature(family_doc) not in existing_doc_sigs
    return family_doc, should_close


class SortableParam(object):
    def __init__(self, fparam):
        self.fparam = fparam

    def __lt__(self, other_sortableparam):
        formula = other_sortableparam.fparam.Formula
        if formula:
            return self.fparam.Definition.Name in formula


def get_symbol_name(symbol_id):
    """Return 'Family : Type' for a valid symbol id, or None if unset."""
    if not symbol_id or symbol_id == DB.ElementId.InvalidElementId:
        return None
    doc = _get_active_doc()
    for fsym in DB.FilteredElementCollector(doc)\
                  .OfClass(DB.FamilySymbol)\
                  .ToElements():
        if fsym.Id == symbol_id:
            return FAMILY_SYMBOL_FORMAT.format(
                revit.query.get_name(fsym.Family),
                revit.query.get_name(fsym)
            )
    return None  # not found

def get_load_class_name(load_class_id):
    """Return 'ELECTRICAL_LOAD_CLASS : <name>' or empty if unset."""
    if not load_class_id or load_class_id == DB.ElementId.InvalidElementId:
        return ELECTRICAL_LOAD_CLASS_FORMAT.format('')
    load_class = _get_active_doc().GetElement(load_class_id)
    name = revit.query.get_name(load_class) if load_class else ''
    return ELECTRICAL_LOAD_CLASS_FORMAT.format(name)


def get_param_typevalue(ftype, fparam):
    '''
    extract value by param type
    '''
    fparam_value = None
    if fparam.StorageType == DB.StorageType.ElementId:
        # Resolve the referenced element id once and guard for "unset"
        eid = ftype.AsElementId(fparam)
        is_unset = (not eid) or (eid == DB.ElementId.InvalidElementId)

        if HOST_APP.is_newer_than(2022):  # ParameterType deprecated in 2023
            dtype = fparam.Definition.GetDataType()

            if DB.Category.IsBuiltInCategory(dtype):
                # Family Type reference (FamilyType param)
                return get_symbol_name(eid) if not is_unset else None

            elif dtype == DB.SpecTypeId.Reference.LoadClassification:
                # Load Classification export even when empty
                return get_load_class_name(eid)  # returns empty formatted string if unset

            else:
                # Other ElementId-backed params store None if unset
                return get_symbol_name(eid) if not is_unset else None

        else:
            # Revit 2022 and earlier
            ptype = fparam.Definition.ParameterType

            if ptype == DB.ParameterType.FamilyType:
                return get_symbol_name(eid) if not is_unset else None

            elif ptype == DB.ParameterType.LoadClassification:
                return get_load_class_name(eid)  # handles unset

            else:
                return get_symbol_name(eid) if not is_unset else None

    elif fparam.StorageType == DB.StorageType.String:
        fparam_value = ftype.AsString(fparam)

#--------------Changed from original to catch unitless integers------------------
    elif fparam.StorageType == DB.StorageType.Integer:
        if hasattr(DB, "SpecTypeId") and hasattr(fparam.Definition, "GetDataType"):
            # Revit 2023+
            spec = fparam.Definition.GetDataType()
            if spec == DB.SpecTypeId.Boolean.YesNo:
                # keep Yes/No as 'true'/'false'
                return 'true' if ftype.AsInteger(fparam) == 1 else 'false'
            else:
                # CATCH-ALL for ALL non-YesNo integer specs (unitless)
                return ftype.AsInteger(fparam)
        else:
            # Revit 2022 and earlier
            if fparam.Definition.ParameterType == DB.ParameterType.YesNo:
                return 'true' if ftype.AsInteger(fparam) == 1 else 'false'
            else:
                # CATCH-ALL for ALL non-YesNo integer specs (unitless)
                return ftype.AsInteger(fparam)
# -------------------------------------------------------------------------------
# -----------------------Handle Doubles------------------------------
    elif fparam.StorageType == DB.StorageType.Double:
        # Revit 2023+ (ForgeTypeId-based)
        if hasattr(DB, "SpecTypeId") and hasattr(fparam.Definition, "GetDataType"):
            spec = fparam.Definition.GetDataType()

            # ---- Unitless double (pure number) ----
            if spec == DB.SpecTypeId.Number:
                return ftype.AsDouble(fparam)

            # ---- Airflow export as numeric CFM (no unit suffix) ----
            elif spec == DB.SpecTypeId.AirFlow:
                try:
                    val_internal = ftype.AsDouble(fparam)
                    return DB.UnitUtils.ConvertFromInternalUnits(
                        val_internal, DB.UnitTypeId.CubicFeetPerMinute
                    )
                except:
                    return ftype.AsValueString(fparam)

            # ---- Everything else: let Revit format it ----
            else:
                return ftype.AsValueString(fparam)

        # Revit 2022 and earlier (ParameterType / DisplayUnitType)
        else:
            ptype = fparam.Definition.ParameterType

            # ---- Unitless double (pure number) ----
            if ptype == DB.ParameterType.Number:
                return ftype.AsDouble(fparam)

            # ---- Airflow export as numeric CFM (no unit suffix) ----
            elif ptype == DB.ParameterType.AirFlow:
                try:
                    val_internal = ftype.AsDouble(fparam)
                    return DB.UnitUtils.ConvertFromInternalUnits(
                        val_internal, DB.DisplayUnitType.DUT_CUBIC_FEET_PER_MINUTE
                    )
                except:
                    return ftype.AsValueString(fparam)

            # ---- Everything else: let Revit format it ----
            else:
                return ftype.AsValueString(fparam)
#-----------------------Handle Doubles------------------------------
    else:
        
        fparam_value = ftype.AsValueString(fparam)

    return fparam_value


def include_type_configs(cfgs_dict, sparams, selected_type_names=None):
    '''
    add the parameter values for all types into the configs dict
    '''
    fm = _get_active_doc().FamilyManager
    selected_names = None
    if selected_type_names:
        selected_names = set([name.strip() for name in selected_type_names if name and name.strip()])
        if not selected_names:
            selected_names = None
    # grab param values for each family type
    for ftype in fm.Types:
        try:
            fname = ftype.Name.strip()
        except Exception:
            fname = ftype.Name
        if selected_names and fname not in selected_names:
            continue
        # param value dict for this type
        type_config = {}
        # grab value from each param
        for sparam in sparams:
            fparam_name = sparam.fparam.Definition.Name
            # add the value to this type config
            type_config[fparam_name] = \
                get_param_typevalue(ftype, sparam.fparam)

        # add the type config to overall config dict
        cfgs_dict[TYPES_SECTION_NAME][ftype.Name] = type_config


def add_default_values(cfgs_dict, sparams):
    '''
    add the parameter values for all types into the configs dict
    '''
    fm = _get_active_doc().FamilyManager
    # grab value from each param
    for sparam in sparams:
        fparam_name = sparam.fparam.Definition.Name
        param_config = cfgs_dict[PARAM_SECTION_NAME][fparam_name]
        # grab current param value
        fparam_value = get_param_typevalue(fm.CurrentType, sparam.fparam)
        if fparam_value:
            param_config[PARAM_SECTION_DEFAULT] = fparam_value


def get_famtype_famcat(fparam):
    '''
    Grab the family category from para with type DB.ParameterType.FamilyType
    These parameters point to a family and symbol but the Revit API
    Does not provide info on what family categories they are assinged to
    '''
    doc = _get_active_doc()
    fm = doc.FamilyManager
    famtype = doc.GetElement(fm.CurrentType.AsElementId(fparam))
    return famtype.Category.Name

#-------------------Added 10-27-25 to put family name at top of yaml------------------

def _get_family_name():
    """Return the family name for this family document."""
    doc = _get_active_doc()
    try:
        fm = doc.FamilyManager
        if hasattr(fm, "OwnerFamily") and fm.OwnerFamily:
            return fm.OwnerFamily.Name
    except Exception:
        pass
    # fallback: file title without .rfa
    title = doc.Title or ""
    return title[:-4] if title.lower().endswith(".rfa") else title

#-------------------Added 10-27-25 to put family name at top of yaml------------------

def read_configs(selected_fparam_names,
                 type_names=None, include_defaults=False):
    '''
    read parameter and type configurations into a dictionary
    '''
    cfgs_dict = dict({
        #------------added 10-27-25
        'family': _get_family_name(),  # <-- NEW: family name at the top
        #------------added 10-27-25
        PARAM_SECTION_NAME: {},
        TYPES_SECTION_NAME: {},
        SHAREDPARAM_DEF: '',
    })

    fm = _get_active_doc().FamilyManager

    # pick the param objects from list of param names
    # params are wrapped by SortableParam
    # SortableParam helps sorting parameters based on their formula
    # dependencies. A parameter that is being used inside another params
    # formula is considered smaller (lower on the output) than that param
    export_sparams = [SortableParam(x) for x in fm.GetParameters()
                      if x.Definition.Name in selected_fparam_names]

    shared_params = []

    # grab all parameter defs
    for sparam in sorted(export_sparams, reverse=True):
        fparam_name = sparam.fparam.Definition.Name
        fparam_isinst = sparam.fparam.IsInstance
        fparam_isreport = sparam.fparam.IsReporting
        fparam_formula = sparam.fparam.Formula
        fparam_shared = sparam.fparam.IsShared
        if HOST_APP.is_newer_than(2022):  # ParameterType deprecated in 2023
            fparam_type = sparam.fparam.Definition.GetDataType()
            fparam_type_str = fparam_type.TypeId
            fparam_group = sparam.fparam.Definition.GetGroupTypeId().TypeId
        else:
            fparam_type = sparam.fparam.Definition.ParameterType
            fparam_type_str = str(fparam_type)
            fparam_group = sparam.fparam.Definition.ParameterGroup

        cfgs_dict[PARAM_SECTION_NAME][fparam_name] = {
            PARAM_SECTION_TYPE: fparam_type_str,
            PARAM_SECTION_GROUP: fparam_group,
            PARAM_SECTION_INST: fparam_isinst,
            PARAM_SECTION_REPORT: fparam_isreport,
            PARAM_SECTION_FORMULA: fparam_formula,
        }

        # add extra data for shared params
        if fparam_shared:
            cfgs_dict[PARAM_SECTION_NAME][fparam_name][PARAM_SECTION_GUID] = \
                sparam.fparam.GUID

        # get the family category if param is FamilyType selector
        if HOST_APP.is_newer_than(2022):  # ParameterType deprecated in 2023
            if 'autodesk.revit.category.family' in fparam_type.TypeId:
                cfgs_dict[PARAM_SECTION_NAME][fparam_name][PARAM_SECTION_CAT] =\
                 get_famtype_famcat(sparam.fparam)
        else:
            if fparam_type == DB.ParameterType.FamilyType:
                cfgs_dict[PARAM_SECTION_NAME][fparam_name]\
                    [PARAM_SECTION_CAT] =\
                    get_famtype_famcat(sparam.fparam)

        # Check if the current family parameter is a shared parameter
        if sparam.fparam.IsShared:
            # Add to an array of sorted shared parameters
            shared_params.append(sparam.fparam)

    # include type configs?
    if type_names:
        include_type_configs(cfgs_dict, export_sparams, type_names)
    elif include_defaults:
        add_default_values(cfgs_dict, export_sparams)

    return cfgs_dict, shared_params


def get_config_file():
    '''
    Get parameter definition yaml file from user
    '''
    save_dir = EXPORT_SOURCE_DIR
    if not os.path.isdir(save_dir):
        try:
            os.makedirs(save_dir)
        except Exception as ex:
            forms.alert(
                'Could not access export folder:\n{}\n\n{}'.format(save_dir, ex),
                title='Export Family Types',
                warn_icon=True
            )
            return None

    selected_path = forms.save_file(
        file_ext='yaml',
        title='Save Family Config YAML',
        init_dir=save_dir,
        default_name=_get_default_export_filename()
    )
    if not selected_path:
        return None

    file_name = os.path.basename(selected_path)
    if not file_name.lower().endswith('.yaml'):
        file_name = file_name + '.yaml'
    forced_path = os.path.join(save_dir, file_name)

    if os.path.normcase(os.path.normpath(selected_path)) != os.path.normcase(os.path.normpath(forced_path)):
        forms.alert(
            'Export location is fixed to:\n{}\n\nSaving as:\n{}'.format(save_dir, forced_path),
            title='Export Family Types'
        )

    return forced_path


def get_parameters():
    '''
    get list of parameters to be exported from user
    '''
    fm = _get_active_doc().FamilyManager
    return forms.SelectFromList.show(
        [x.Definition.Name for x in fm.GetParameters()],
        title="Select Parameters",
        multiselect=True,
    ) or []


def get_type_names_for_export():
    '''
    Prompt the user to select specific family types to export.
    Returns:
        None -> user cancelled the selection
        [] -> user chose to skip exporting types
        [names] -> selected family type names
    '''
    fm = _get_active_doc().FamilyManager
    type_names = []
    for ftype in fm.Types:
        try:
            name = ftype.Name.strip()
        except Exception:
            name = ftype.Name
        if name:
            type_names.append(name)

    if not type_names:
        return []

    type_names = sorted(set(type_names), key=lambda n: n.lower())
    selected = forms.SelectFromList.show(
        type_names,
        title='Select Family Types to Export',
        button_name='Use Selected Types',
        multiselect=True,
        message='Choose one or more family types to include.\n'
                'Leave empty to export parameters only.'
    )
    if selected is None:
        return None

    cleaned = [name.strip() for name in selected if name and name.strip()]
    return cleaned


def store_sharedparam_def(shared_params):
    '''
    Reads the shared parameters into a txt file
    '''
    sparam_file = HOST_APP.app.OpenSharedParameterFile()
    exported_sparams_grp = sparam_file.Groups.Create("Exported Parameters")
    for sparam in shared_params:
        if HOST_APP.is_newer_than(2022):  # ParameterType deprecated in 2023
            param_type = sparam.Definition.GetDataType()
        else:
            param_type = sparam.Definition.ParameterType
        sparamdef_create_options = \
            DB.ExternalDefinitionCreationOptions(
                sparam.Definition.Name,
                param_type,
                GUID=sparam.GUID
            )

        try:
            exported_sparams_grp.Definitions.Create(sparamdef_create_options)
        except Exceptions.ArgumentException:
            forms.alert("A parameter with the same GUID already exists.\nParameter: {} will be ignored.".format(sparam.Definition.Name))


def get_shared_param_def_contents(shared_params):
    '''
    get a temporary text file to store the generated shared param data
    '''
    global family_cfg_file
    temp_defs_filepath = \
        script.get_instance_data_file(
            file_id=coreutils.get_file_name(family_cfg_file),
            add_cmd_name=True
        )
    # make sure the ParameterGroup file exists and it is empty
    open(temp_defs_filepath, 'wb').close()
    # swap existing shared param with temp
    existing_sharedparam_file = HOST_APP.app.SharedParametersFilename
    HOST_APP.app.SharedParametersFilename = temp_defs_filepath
    # write the shared param data
    store_sharedparam_def(shared_params)
    # restore the original shared param file
    HOST_APP.app.SharedParametersFilename = existing_sharedparam_file

    return revit.files.read_text(temp_defs_filepath)


def save_configs(configs_dict, param_file):
    '''
    Load contents of yaml file into an ordered dict
    '''
    return yaml.dump_dict(configs_dict, param_file)


def _get_map_file(export_yaml_file):
    return os.path.join(FAMILY_CONFIGS_ROOT, MAP_FILE_NAME)


def _get_exported_type_names(family_configs):
    type_names = []
    types_cfg = family_configs.get(TYPES_SECTION_NAME, {}) or {}
    for type_name in types_cfg.keys():
        if not type_name:
            continue
        cleaned_name = type_name.strip()
        if cleaned_name:
            type_names.append(cleaned_name)
    return sorted(set(type_names), key=lambda n: n.lower())


def _update_map_file(export_yaml_file, family_configs):
    map_file = _get_map_file(export_yaml_file)
    map_dir = os.path.dirname(map_file)
    if map_dir and not os.path.isdir(map_dir):
        os.makedirs(map_dir)
    map_data = {}

    if os.path.exists(map_file):
        try:
            loaded_map = yaml.load_as_dict(map_file)
            if isinstance(loaded_map, dict):
                map_data = loaded_map
            else:
                logger.warning('Existing map.yaml is not a dictionary. Rebuilding map file.')
        except Exception as load_error:
            logger.warning('Could not read existing map.yaml. Rebuilding map file. | %s', load_error)

    map_key = os.path.basename(export_yaml_file)
    map_data[map_key] = _get_exported_type_names(family_configs)
    yaml.dump_dict(map_data, map_file)
    return map_file


if __name__ == '__main__':
    family_doc = None
    should_close_family_doc = False
    previous_active_doc = ACTIVE_DOC

    try:
        family_doc, should_close_family_doc = _resolve_family_doc_for_export()
        if not family_doc:
            raise SystemExit

        ACTIVE_DOC = family_doc

        family_cfg_file = get_config_file()
        if family_cfg_file:
            family_params = get_parameters()
            if family_params:
                selected_type_names = get_type_names_for_export()
                if selected_type_names is None:
                    forms.alert('Export cancelled.', warn_icon=True)
                    raise SystemExit

                include_defaults = False
                if not selected_type_names:
                    include_defaults = forms.alert(
                        "Do you want to include the current parameter values as "
                        "default? Otherwise the parameters will not include any "
                        "value and their default value will be assigned "
                        "by Revit at import.",
                        yes=True, no=True
                    )

                family_configs, shared_parameters = \
                    read_configs(family_params,
                                 type_names=selected_type_names,
                                 include_defaults=include_defaults)

                logger.debug(family_configs)

                # get revit to generate contents of a shared param file definition
                # for the shared parameters and store that inside the yaml file
                if shared_parameters:
                    family_configs[SHAREDPARAM_DEF] = \
                        get_shared_param_def_contents(shared_parameters)

                save_configs(family_configs, family_cfg_file)
                _update_map_file(family_cfg_file, family_configs)
    finally:
        ACTIVE_DOC = previous_active_doc
        if family_doc and should_close_family_doc:
            try:
                family_doc.Close(False)
            except Exception:
                pass
