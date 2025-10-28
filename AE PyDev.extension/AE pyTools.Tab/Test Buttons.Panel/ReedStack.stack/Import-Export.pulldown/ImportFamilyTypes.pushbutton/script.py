# -*- coding: utf-8 -*-
"""Import family configurations (including shared parameters) from YAML.

Now supports PROJECT environment:
- If in a project document:
  * Pick a loaded Family by name
  * EditFamily() opens a temporary family document
  * Import parameters + selected types into that family
  * LoadFamily() back into the project (auto-overwrite handler)
  * Close the temp family doc

Performance-minded:
- Avoids inserting formulas entirely
- Skips redundant value sets by comparing current values
- Caches parameter lookups and symbol/load-class resolutions
- (New) Skip instance parameters entirely for faster runs
"""
from __future__ import print_function

from collections import namedtuple

from pyrevit import coreutils
from pyrevit import revit, DB, HOST_APP
from pyrevit import forms
from pyrevit import script
from pyrevit.coreutils import yaml

logger = script.get_logger()
output = script.get_output()

# ---- Import mode toggles ----------------------------------------------------
SKIP_INSTANCE_PARAMS = True  # only import TYPE parameters (faster)
# -----------------------------------------------------------------------------

# ----------------------------- YAML keys -------------------------------------
PARAM_SECTION_NAME = 'parameters'
PARAM_SECTION_TYPE = 'type'
PARAM_SECTION_CAT = 'category'
PARAM_SECTION_GROUP = 'group'
PARAM_SECTION_INST = 'instance'
PARAM_SECTION_REPORT = 'reporting'
PARAM_SECTION_FORMULA = 'formula'
PARAM_SECTION_DEFAULT = 'default'
PARAM_SECTION_GUID = 'GUID'

TYPES_SECTION_NAME = 'types'
SHAREDPARAM_DEF = 'xref_sharedparams'
# -----------------------------------------------------------------------------

DEFAULT_TYPE = 'Text'
if HOST_APP.is_newer_than(2022):
    DEFAULT_PARAM_GROUP = 'Construction'
else:
    DEFAULT_PARAM_GROUP = 'PG_CONSTRUCTION'

FAMILY_SYMBOL_SEPARATOR = ' : '
TEMP_TYPENAME = "Default"

ParamConfig = namedtuple(
    'ParamConfig',
    ['name', 'bigroup', 'bitype', 'famcat',
     'isinst', 'isreport', 'formula', 'default', 'GUID']
)
ParamValueConfig = namedtuple('ParamValueConfig', ['name', 'value'])
TypeConfig = namedtuple('TypeConfig', ['name', 'param_values'])

failed_params = []

# ------------------------- Project-mode helpers ------------------------------

class _AlwaysOverwriteLoader(DB.IFamilyLoadOptions):
    def __init__(self, overwrite_params=True):
        self._overwrite_params = overwrite_params
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = self._overwrite_params
        return True
    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Set(DB.FamilySource.Family)
        overwriteParameterValues.Value = self._overwrite_params
        return True

def _pick_loaded_family(project_doc):
    fams = list(DB.FilteredElementCollector(project_doc).OfClass(DB.Family))
    if not fams:
        forms.alert('No loaded Families found in this project.', warn_icon=True)
        return None
    items = sorted([(f, f.Name) for f in fams], key=lambda x: x[1].lower())
    choice = forms.SelectFromList.show(
        [name for _, name in items],
        title='Select Family to Edit/Update',
        multiselect=False,
        button_name='Edit Family'
    )
    if not choice:
        return None
    for f, name in items:
        if name == choice:
            return f
    return None

def _get_yaml_family_name(family_configs):
    """Read 'family' from YAML (string or None)."""
    name = family_configs.get('family')
    return name.strip() if isinstance(name, basestring) and name.strip() else None


def _find_family_by_name(project_doc, name):
    """Return a DB.Family by exact name match, else None."""
    if not name:
        return None
    for f in DB.FilteredElementCollector(project_doc).OfClass(DB.Family):
        try:
            if f.Name.strip() == name.strip():
                return f
        except Exception:
            pass
    return None

# ------------------------- Caches --------------------------------------------

_SYMBOL_ID_CACHE = {}          # (fam_name, sym_name) -> ElementId
_LOAD_CLASS_ID_CACHE = {}      # load_class_name -> ElementId

def parse_familysymbol_refvalue(param_value):
    if FAMILY_SYMBOL_SEPARATOR not in param_value:
        logger.warning(
            'Family type parameter value must be formatted as '
            '<family-name> : <symbol-name> | incorrect: %s', param_value
        )
        return
    return param_value.split(FAMILY_SYMBOL_SEPARATOR, 1)

def _cache_key_symbol(fam_name, sym_name):
    return (fam_name or '').strip(), (sym_name or '').strip()

def get_symbol_id(doc, param_value):
    fam_name, sym_name = parse_familysymbol_refvalue(param_value)
    if fam_name is None:
        return None
    key = _cache_key_symbol(fam_name, sym_name)
    if key in _SYMBOL_ID_CACHE:
        return _SYMBOL_ID_CACHE[key]
    for fsym in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol):
        try:
            famname = fsym.Family.Name if fsym.Family else None
            symname = fsym.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            if famname == fam_name and symname == sym_name:
                _SYMBOL_ID_CACHE[key] = fsym.Id
                return fsym.Id
        except Exception:
            pass
    _SYMBOL_ID_CACHE[key] = None
    return None

def get_load_class_id(doc, param_value):
    load_class, load_class_name = parse_familysymbol_refvalue(param_value)
    if load_class != "ELECTRICAL_LOAD_CLASS":
        return None
    name = (load_class_name or '').strip()
    if name in _LOAD_CLASS_ID_CACHE:
        return _LOAD_CLASS_ID_CACHE[name]
    for lc in DB.FilteredElementCollector(doc).OfClass(DB.Electrical.ElectricalLoadClassification):
        if lc.Name == name:
            _LOAD_CLASS_ID_CACHE[name] = lc.Id
            return lc.Id
    new_lc = DB.Electrical.ElectricalLoadClassification.Create(doc, name)
    if new_lc:
        _LOAD_CLASS_ID_CACHE[name] = new_lc.Id
        return new_lc.Id
    return None

# ------------------------- Param config helpers ------------------------------

def _get_category_by_name(doc, name_or_bic):
    # try BuiltInCategory enum by string
    try:
        bic = getattr(DB.BuiltInCategory, name_or_bic)
        return doc.Settings.Categories.get_Item(bic)
    except Exception:
        pass
    # fallback: display name
    for cat in doc.Settings.Categories:
        if cat.Name == name_or_bic:
            return cat
    return None

def get_param_config(doc, param_name, param_opts):
    if HOST_APP.is_newer_than(2022):
        param_group = DB.ForgeTypeId(param_opts.get(PARAM_SECTION_GROUP, DEFAULT_PARAM_GROUP))
        param_famtype = None
        param_type = DB.ForgeTypeId(param_opts.get(PARAM_SECTION_TYPE, DEFAULT_TYPE))
        if param_opts.get(PARAM_SECTION_CAT):
            param_famtype = _get_category_by_name(doc, param_opts.get(PARAM_SECTION_CAT))
    else:
        param_group = coreutils.get_enum_value(
            DB.BuiltInParameterGroup,
            param_opts.get(PARAM_SECTION_GROUP, DEFAULT_PARAM_GROUP)
        )
        param_type = coreutils.get_enum_value(
            DB.ParameterType,
            param_opts.get(PARAM_SECTION_TYPE, DEFAULT_TYPE)
        )
        param_famtype = None
        if param_type == DB.ParameterType.FamilyType and param_opts.get(PARAM_SECTION_CAT):
            param_famtype = _get_category_by_name(doc, param_opts.get(PARAM_SECTION_CAT))

    param_isinst = str(param_opts.get(PARAM_SECTION_INST, 'false')).lower() == 'true'
    param_isreport = str(param_opts.get(PARAM_SECTION_REPORT, 'false')).lower() == 'true'
    param_formula = param_opts.get(PARAM_SECTION_FORMULA, None)  # read but never apply
    param_default = param_opts.get(PARAM_SECTION_DEFAULT, None)
    param_GUID = param_opts.get(PARAM_SECTION_GUID, None)

    if not param_group or not param_type:
        logger.critical('Can not determine group/type for %s', param_name)
        return None

    return ParamConfig(
        name=param_name,
        bigroup=param_group,
        bitype=param_type,
        famcat=param_famtype,
        isinst=param_isinst,
        isreport=param_isreport,
        formula=param_formula,
        default=param_default,
        GUID=param_GUID
    )

# ------------------------- Value comparison helpers --------------------------

def _normalize_bool_string(v):
    s = str(v).strip().lower()
    if s in ('true', '1', 'yes'):
        return 'true'
    if s in ('false', '0', 'no'):
        return 'false'
    return s

def _current_value_string(ftype, fparam):
    """Return current value for the active FamilyType as a comparable string."""
    try:
        st = fparam.StorageType
        if st == DB.StorageType.Integer:
            # account for Yes/No
            if HOST_APP.is_newer_than(2022):
                if fparam.Definition.GetDataType() == DB.SpecTypeId.Boolean.YesNo:
                    return 'true' if ftype.AsInteger(fparam) == 1 else 'false'
            else:
                if getattr(DB.ParameterType, 'YesNo', None) and fparam.Definition.ParameterType == DB.ParameterType.YesNo:
                    return 'true' if ftype.AsInteger(fparam) == 1 else 'false'
            return str(ftype.AsInteger(fparam))
        elif st == DB.StorageType.String:
            return ftype.AsString(fparam) or ''
        elif st == DB.StorageType.Double:
            # use Revit's formatted value for apples-to-apples compare
            return ftype.AsValueString(fparam) or ''
        elif st == DB.StorageType.ElementId:
            # comparing IDs is messy across YAML; skip compare so we always set
            return None
        else:
            return ftype.AsValueString(fparam) or ''
    except Exception:
        return None

def _incoming_value_string(pvcfg, fparam):
    """Normalize incoming YAML value to a comparable string for same storage type."""
    v = pvcfg.value
    if v is None:
        return None
    try:
        st = fparam.StorageType
        if st == DB.StorageType.Integer:
            if HOST_APP.is_newer_than(2022):
                if fparam.Definition.GetDataType() == DB.SpecTypeId.Boolean.YesNo:
                    return _normalize_bool_string(v)
            else:
                if getattr(DB.ParameterType, 'YesNo', None) and fparam.Definition.ParameterType == DB.ParameterType.YesNo:
                    return _normalize_bool_string(v)
            return str(int(v))
        elif st == DB.StorageType.String:
            return '' if v is None else str(v)
        elif st == DB.StorageType.Double:
            # incoming might be "120 V" or "3'-0\""—compare against formatted string path
            return str(v)
        elif st == DB.StorageType.ElementId:
            return None
        else:
            return str(v)
    except Exception:
        return None

def _values_equal(ftype, fparam, pvcfg):
    cur = _current_value_string(ftype, fparam)
    inc = _incoming_value_string(pvcfg, fparam)
    if cur is None or inc is None:
        return False
    return cur == inc

# ------------------------- Setters -------------------------------------------

def set_fparam_value(doc, pvcfg, fparam):
    fm = doc.FamilyManager
    if fparam.Formula:
        logger.debug('parameter has an existing formula; skipping set: %s', pvcfg.name)
        return
    if pvcfg.value in (None, ''):
        logger.debug('skipping parameter with no value: %s', pvcfg.name)
        return

    # ElementId-like params (family type refs, classifications) → no pre-compare
    if fparam.StorageType == DB.StorageType.ElementId:
        if HOST_APP.is_newer_than(2022):
            if DB.Category.IsBuiltInCategory(fparam.Definition.GetDataType()):
                fsym_id = get_symbol_id(doc, pvcfg.value)
                if fsym_id: fm.Set(fparam, fsym_id); return
            if fparam.Definition.GetDataType() == DB.SpecTypeId.Reference.LoadClassification:
                lc_id = get_load_class_id(doc, pvcfg.value)
                if lc_id: fm.Set(fparam, lc_id); return
        else:
            if fparam.Definition.ParameterType == DB.ParameterType.FamilyType:
                fsym_id = get_symbol_id(doc, pvcfg.value)
                if fsym_id: fm.Set(fparam, fsym_id); return
            if fparam.Definition.ParameterType == DB.ParameterType.LoadClassification:
                lc_id = get_load_class_id(doc, pvcfg.value)
                if lc_id: fm.Set(fparam, lc_id); return
        return

    # Compare current vs incoming; skip if equal
    ftype = fm.CurrentType
    try:
        if _values_equal(ftype, fparam, pvcfg):
            return
    except Exception:
        pass

    st = fparam.StorageType
    if st == DB.StorageType.String:
        fm.Set(fparam, pvcfg.value)
    elif st == DB.StorageType.Integer:
        if HOST_APP.is_newer_than(2022):
            if DB.SpecTypeId.Boolean.YesNo == fparam.Definition.GetDataType():
                fm.Set(fparam, 1 if _normalize_bool_string(pvcfg.value) == 'true' else 0)
            else:
                fm.Set(fparam, int(pvcfg.value))
        else:
            if getattr(DB.ParameterType, 'YesNo', None) and fparam.Definition.ParameterType == DB.ParameterType.YesNo:
                fm.Set(fparam, 1 if _normalize_bool_string(pvcfg.value) == 'true' else 0)
            else:
                fm.Set(fparam, int(pvcfg.value))
    else:
        # rely on Revit to parse e.g., "120 V", "3'-0\""
        fm.SetValueString(fparam, str(pvcfg.value))

# ------------------------- Ensure (create + set) -----------------------------

def ensure_param_value(doc, fm, fparam, pcfg, param_name):
    # Skip instance parameters entirely if toggle enabled
    if SKIP_INSTANCE_PARAMS and fparam.IsInstance:
        return

    # Ignore formulas entirely (do not insert)
    # Apply default only if it would change something
    if pcfg.default is not None and not fparam.IsReporting:
        try:
            pvcfg = ParamValueConfig(name=pcfg.name, value=pcfg.default)
            ftype = fm.CurrentType
            if not _values_equal(ftype, fparam, pvcfg):
                set_fparam_value(doc, pvcfg, fparam)
        except Exception as ex:
            logger.error('Failed to set default for %s | %s', pcfg.name, ex)

    if pcfg.isreport:
        try:
            fm.MakeReporting(fparam)
        except Exception as ex:
            logger.error('Failed to make reporting: %s | %s', pcfg.name, ex)

def _get_param_by_name_map(doc):
    """Build a one-shot map of parameter name -> FamilyParameter."""
    fmap = {}
    for fp in doc.FamilyManager.Parameters:
        try:
            fmap[fp.Definition.Name] = fp
        except Exception:
            pass
    return fmap

def ensure_param(doc, fm, pcfg, param_name):
    # Skip creating instance parameters entirely if toggle enabled
    if SKIP_INSTANCE_PARAMS and pcfg.isinst:
        logger.debug('Skipping instance parameter by mode: %s', param_name)
        return None

    def _get_family_parameter(_doc, name):
        for fp in _doc.FamilyManager.Parameters:
            if fp.Definition.Name == name:
                return fp
        return None

    fparam = _get_family_parameter(doc, param_name)
    if fparam:
        return fparam

    try:
        # Shared parameter by GUID?
        if pcfg.GUID is not None:
            sparam_found = False
            sparam_file = HOST_APP.app.OpenSharedParameterFile()
            if sparam_file:
                for def_grp in sparam_file.Groups:
                    if def_grp.Name == "Exported Parameters":
                        for sparam_def in def_grp.Definitions:
                            if str(sparam_def.GUID) == pcfg.GUID:
                                sparam_found = True
                                fparam = fm.AddParameter(sparam_def, pcfg.bigroup, pcfg.isinst)
                                break
            if not sparam_found:
                logger.error('Shared parameter definition not found for %s', param_name)
                return None
        else:
            fparam = fm.AddParameter(
                pcfg.name,
                pcfg.bigroup,
                pcfg.famcat if pcfg.famcat else pcfg.bitype,
                pcfg.isinst
            )
    except Exception as ex:
        failed_params.append(pcfg.name)
        if pcfg.famcat:
            logger.error(
                'Error creating parameter: %s\n'
                'Likely a nested family selector—ensure at least one nested family '
                'of type "%s" is loaded. | %s', pcfg.name, pcfg.famcat.Name, ex
            )
        else:
            logger.error('Error creating parameter: %s | %s', pcfg.name, ex)
        return None

    return fparam

def ensure_params(doc, fconfig):
    params_with_value = []
    param_cfgs = fconfig.get(PARAM_SECTION_NAME, None)
    fm = doc.FamilyManager
    if not param_cfgs:
        return

    for pname, popts in param_cfgs.items():
        if not (pname and popts):
            continue
        pcfg = get_param_config(doc, pname, popts)
        if not pcfg:
            continue

        # Skip instance parameters entirely if toggle enabled
        if SKIP_INSTANCE_PARAMS and pcfg.isinst:
            logger.debug('Skipping instance parameter by mode: %s', pname)
            continue

        fparam = ensure_param(doc, fm, pcfg, pname)
        if fparam and (pcfg.default is not None or pcfg.isreport):
            params_with_value.append((fparam, pcfg, pname))

    # single pass for defaults/reporting
    for fparam, pcfg, pname in params_with_value:
        ensure_param_value(doc, fm, fparam, pcfg, pname)

def get_type_config(type_name, type_opts):
    if type_name and type_opts:
        pvalue_cfgs = []
        for pname, pvalue in type_opts.items():
            pvalue_cfgs.append(ParamValueConfig(name=pname, value=pvalue))
        return TypeConfig(name=type_name, param_values=pvalue_cfgs)

def ensure_type(doc, type_config):
    fm = doc.FamilyManager
    for t in fm.Types:
        if t.Name == type_config.name:
            return t
    return fm.NewType(type_config.name)

def ensure_types(doc, fconfig):
    fm = doc.FamilyManager
    type_cfgs = fconfig.get(TYPES_SECTION_NAME, None)
    if not type_cfgs:
        return
    # build parameter name map once
    param_map = _get_param_by_name_map(doc)

    for tname, topts in type_cfgs.items():
        tname = tname.strip()
        tcfg = get_type_config(tname, topts)
        if not tcfg:
            continue
        ftype = ensure_type(doc, tcfg)
        if not ftype:
            continue

        fm.CurrentType = ftype
        for pvcfg in tcfg.param_values:
            fparam = param_map.get(pvcfg.name)
            if not fparam:
                logger.debug('param not found on family: %s', pvcfg.name)
                continue

            # Skip instance params entirely if toggle enabled
            if SKIP_INSTANCE_PARAMS and fparam.IsInstance:
                logger.debug('Skipping instance param on type set: %s', pvcfg.name)
                continue

            if fparam.IsReporting:
                logger.debug('skip reporting param: %s', pvcfg.name)
                continue

            # skip redundant write
            try:
                if _values_equal(fm.CurrentType, fparam, pvcfg):
                    continue
            except Exception:
                pass

            set_fparam_value(doc, pvcfg, fparam)

# ------------------------- YAML & Type Picker --------------------------------

def get_config_file():
    return forms.pick_file(file_ext='yaml')

def load_configs(parma_file):
    return yaml.load_as_dict(parma_file)

def recover_sharedparam_defs(family_cfg_file, sharedparam_def_contents):
    temp_defs_filepath = script.get_instance_data_file(
        file_id=coreutils.get_file_name(family_cfg_file),
        add_cmd_name=True
    )
    revit.files.write_text(temp_defs_filepath, sharedparam_def_contents)
    return temp_defs_filepath

def pick_types_from_yaml(family_configs):
    types_dict = family_configs.get(TYPES_SECTION_NAME, {}) or {}
    all_type_names = sorted([t.strip() for t in types_dict.keys()])

    if not all_type_names:
        forms.alert('No types were found in the YAML.', title='Type Picker', warn_icon=True)
        return []

    selected = forms.SelectFromList.show(
        all_type_names,
        title='Select Family Types to Import',
        button_name='Load Selected Types',
        multiselect=True
    )

    if selected is None:
        return None

    if len(selected) > 50:
        res = forms.alert(
            'You selected {} types.\nLarge imports can be very slow.\n\n'
            'Continue anyway?'.format(len(selected)),
            title='Confirm Large Import',
            warn_icon=True,
            yes=True, no=True
        )
        if not res:
            return None

    return selected

def filtered_config_for_types(family_configs, selected_type_names):
    if selected_type_names is None:
        return None
    new_cfg = {}
    if PARAM_SECTION_NAME in family_configs:
        new_cfg[PARAM_SECTION_NAME] = family_configs[PARAM_SECTION_NAME]
    if SHAREDPARAM_DEF in family_configs:
        new_cfg[SHAREDPARAM_DEF] = family_configs[SHAREDPARAM_DEF]
    all_types = family_configs.get(TYPES_SECTION_NAME, {}) or {}
    new_cfg[TYPES_SECTION_NAME] = {
        k: all_types[k] for k in all_types.keys() if k.strip() in set(selected_type_names)
    }
    return new_cfg

# --------------------------------- MAIN --------------------------------------

if __name__ == '__main__':
    family_cfg_file = get_config_file()
    if not family_cfg_file:
        raise SystemExit

    family_configs = load_configs(family_cfg_file)
    existing_sharedparam_file = HOST_APP.app.SharedParametersFilename

    selected_type_names = pick_types_from_yaml(family_configs)
    if selected_type_names is None:
        forms.alert('Import cancelled.', title='Import Family Types')
        raise SystemExit
    filtered_configs = filtered_config_for_types(family_configs, selected_type_names)

    project_doc = revit.doc
    is_family_doc = project_doc.IsFamilyDocument

    fam_doc = None
    loader = _AlwaysOverwriteLoader(overwrite_params=True)

    try:
        yaml_family_name = _get_yaml_family_name(family_configs)

        if is_family_doc:
            # Use the currently open family; optionally warn if mismatch
            fam_doc = project_doc
            try:
                current_name = fam_doc.FamilyManager.OwnerFamily.Name
            except Exception:
                current_name = (fam_doc.Title[:-4] if fam_doc.Title.lower().endswith('.rfa') else fam_doc.Title)
            if yaml_family_name and current_name.strip() != yaml_family_name:
                logger.warning("YAML family '%s' does not match open family '%s'. Proceeding anyway.",
                               yaml_family_name, current_name)
        else:
            # Auto-find the family in the project by name from YAML (no prompt)
            if not yaml_family_name:
                forms.alert(
                    "This config has no 'family' key at the top, so I can't auto-select.\n"
                    "Re-export with the updated exporter (it writes family: <Name>).",
                    warn_icon=True, title='Import Family Types'
                )
                raise SystemExit

            target_family = _find_family_by_name(project_doc, yaml_family_name)
            if not target_family:
                forms.alert(
                    'Family "{}" not found in this project. Load it, then rerun import.'.format(yaml_family_name),
                    warn_icon=True, title='Import Family Types'
                )
                raise SystemExit

            # Open a temporary editable family document
            fam_doc = project_doc.EditFamily(target_family)

        if SHAREDPARAM_DEF in family_configs:
            sharedparam_file = recover_sharedparam_defs(family_cfg_file, family_configs[SHAREDPARAM_DEF])
            HOST_APP.app.SharedParametersFilename = sharedparam_file

        with revit.Transaction('Import Params/Types from Config', doc=fam_doc):
            fam_mgr = fam_doc.FamilyManager
            ctype = fam_mgr.CurrentType or fam_mgr.NewType(TEMP_TYPENAME)

            ensure_params(fam_doc, family_configs)
            ensure_types(fam_doc, filtered_configs)

            if ctype and ctype.Name != TEMP_TYPENAME:
                fam_mgr.CurrentType = ctype

        if not is_family_doc:
            fam_doc.LoadFamily(project_doc, loader)
            forms.alert('Family updated in project.', title='Import Family Types')

    except Exception as import_error:
        logger.error('Import failed: %s', import_error)
        raise
    finally:
        HOST_APP.app.SharedParametersFilename = existing_sharedparam_file
        if fam_doc and (fam_doc != project_doc):
            try:
                fam_doc.Close(False)
            except Exception:
                pass
