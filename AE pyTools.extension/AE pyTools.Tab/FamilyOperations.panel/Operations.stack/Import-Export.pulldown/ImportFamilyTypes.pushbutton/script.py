# -*- coding: utf-8 -*-
"""Import family configurations (including shared parameters) from YAML.

Now supports map-driven PROJECT imports:
- User selects:
  * map.yaml (lookup index)
  * target family .rfa
- Import resolves the matching family YAML from map.yaml, then:
  * prompts for type selection
  * imports parameters + selected types into the family document
  * loads the family back into the project (new or existing)

Performance-minded:
- Avoids inserting formulas entirely
- Skips redundant value sets by comparing current values
- Caches parameter lookups and symbol/load-class resolutions
- (New) Skip instance parameters entirely for faster runs
"""
from __future__ import print_function

from collections import namedtuple
import os
import shutil

from pyrevit import coreutils
from pyrevit import revit, DB, HOST_APP
from pyrevit import forms
from pyrevit import script
from pyrevit.coreutils import yaml

try:
    from System.Windows import Visibility
except Exception:
    class _VisibilityShim(object):
        Visible = None
        Collapsed = None
    Visibility = _VisibilityShim()

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
MAP_FILE_NAME = 'map.yaml'
FAMILY_CONFIGS_ROOT = r'C:\ACC\ACCDocs\CoolSys\CED Content Collection\Project Files\03 Automations\Family Configs'
RFA_PICKER_DEFAULT_DIR = r'C:\ACC\ACCDocs\CoolSys\_CED Revit Development Project\Project Files\00 Development Files\01 Family Libraries\01 CED Standard Content'
# -----------------------------------------------------------------------------

DEFAULT_TYPE = 'Text'
if HOST_APP.is_newer_than(2022):
    DEFAULT_PARAM_GROUP = 'Construction'
else:
    DEFAULT_PARAM_GROUP = 'PG_CONSTRUCTION'

FAMILY_SYMBOL_SEPARATOR = ' : '
TEMP_TYPENAME = "Default"
RESOLUTION_WINDOW_XAML = os.path.join(os.path.dirname(__file__), 'resolution_window.xaml')

ParamConfig = namedtuple(
    'ParamConfig',
    ['name', 'bigroup', 'bitype', 'famcat',
     'isinst', 'isreport', 'formula', 'default', 'GUID']
)
ParamValueConfig = namedtuple('ParamValueConfig', ['name', 'value'])
TypeConfig = namedtuple('TypeConfig', ['name', 'param_values'])

failed_params = []

class ImportAlertTracker(object):
    def __init__(self):
        self.reset()

    def reset(self):
        self.formula_conflicts = {}      # param -> {'yaml': str, 'family': str}
        self.yaml_formula_only = {}      # param -> yaml formula
        self.family_formula_only = {}    # param -> family formula
        self.missing_parameters = set()  # YAML params absent in family
        self.extra_family_parameters = set()  # Family params absent in YAML
        self.instance_mismatches = []    # [{parameter, yaml_isinst, family_isinst}]
        self._reported = False

    def add_formula_conflict(self, param_name, yaml_formula, family_formula):
        self.formula_conflicts[param_name] = {
            'yaml': yaml_formula,
            'family': family_formula
        }

    def add_yaml_formula_only(self, param_name, yaml_formula):
        self.yaml_formula_only[param_name] = yaml_formula

    def add_family_formula_only(self, param_name, family_formula):
        self.family_formula_only[param_name] = family_formula

    def add_missing_parameter(self, param_name):
        self.missing_parameters.add(param_name)

    def add_extra_family_parameter(self, param_name):
        self.extra_family_parameters.add(param_name)

    def add_instance_mismatch(self, param_name, yaml_is_instance, family_is_instance):
        self.instance_mismatches.append({
            'parameter': param_name,
            'yaml_isinst': bool(yaml_is_instance),
            'family_isinst': bool(family_is_instance)
        })

    def has_alerts(self):
        return bool(
            self.formula_conflicts
            or self.yaml_formula_only
            or self.family_formula_only
            or self.missing_parameters
            or self.extra_family_parameters
            or self.instance_mismatches
        )

    def has_actionable_alerts(self):
        return bool(
            self.formula_conflicts
            or self.yaml_formula_only
            or self.family_formula_only
            or self.missing_parameters
        )

    def get_formula_category(self, param_name):
        if param_name in self.formula_conflicts:
            return 'conflict'
        if param_name in self.yaml_formula_only:
            return 'yaml_only'
        if param_name in self.family_formula_only:
            return 'family_only'
        return None

    def is_missing_parameter(self, param_name):
        return param_name in self.missing_parameters

    def report(self):
        if self._reported or not self.has_alerts():
            return True

        self._reported = True
        output.print_md('### Parameter Alerts')

        if self.formula_conflicts:
            output.print_md('**Formula mismatches ({})**'.format(len(self.formula_conflicts)))
            for pname, payload in sorted(self.formula_conflicts.items()):
                output.print_md(
                    '- {param} -> YAML: {yaml} | Family: {family}'.format(
                        param=self._format_inline(pname),
                        yaml=self._format_inline(payload.get('yaml')),
                        family=self._format_inline(payload.get('family'))
                    )
                )

        if self.yaml_formula_only:
            output.print_md('**YAML-only formulas ({})**'.format(len(self.yaml_formula_only)))
            for pname, yformula in sorted(self.yaml_formula_only.items()):
                output.print_md(
                    '- {param} -> YAML formula: {yaml}'.format(
                        param=self._format_inline(pname),
                        yaml=self._format_inline(yformula)
                    )
                )

        if self.family_formula_only:
            output.print_md('**Family-only formulas ({})**'.format(len(self.family_formula_only)))
            for pname, fformula in sorted(self.family_formula_only.items()):
                output.print_md(
                    '- {param} -> Family formula: {family}'.format(
                        param=self._format_inline(pname),
                        family=self._format_inline(fformula)
                    )
                )

        if self.missing_parameters:
            output.print_md('**YAML parameters missing from family ({})**'.format(len(self.missing_parameters)))
            for pname in sorted(self.missing_parameters):
                output.print_md('- {}'.format(self._format_inline(pname)))

        if self.extra_family_parameters:
            output.print_md('**Family parameters not present in YAML ({})**'.format(len(self.extra_family_parameters)))
            for pname in sorted(self.extra_family_parameters):
                output.print_md('- {}'.format(self._format_inline(pname)))

        if self.instance_mismatches:
            output.print_md('**Instance vs Type mismatches ({})**'.format(len(self.instance_mismatches)))
            for alert in self.instance_mismatches:
                output.print_md(
                    '- {param} -> YAML: {yaml} | Family: {family}'.format(
                        param=self._format_inline(alert['parameter']),
                        yaml='Instance' if alert['yaml_isinst'] else 'Type',
                        family='Instance' if alert['family_isinst'] else 'Type'
                    )
                )

        forms.alert(
            'Parameter discrepancies detected. See output panel for detailed lists.',
            title='Import Parameter Alerts',
            warn_icon=True
        )
        return True

    @staticmethod
    def _format_inline(value):
        if value in (None, ''):
            safe_value = '(none)'
        else:
            safe_value = str(value)
        safe_value = safe_value.replace('`', '\\`')
        return '`{}`'.format(safe_value)


class ImportResolutionOptions(object):
    ACTION_USE_YAML_FORMULA = 'use_yaml_formula'
    ACTION_KEEP_TARGET_FORMULA = 'keep_target_formula'
    ACTION_CLEAR_AND_USE_VALUE = 'clear_formula_use_value'
    ACTION_USE_YAML_VALUE_ONLY = 'use_yaml_value_only'

    ACTION_ADD_WITH_FORMULA = 'add_with_formula'
    ACTION_ADD_WITH_VALUE = 'add_with_value'
    ACTION_IGNORE_MISSING = 'ignore_missing'

    def __init__(self):
        self.reset()

    def reset(self):
        self.formula_conflict_action = self.ACTION_KEEP_TARGET_FORMULA
        self.yaml_formula_only_action = self.ACTION_USE_YAML_VALUE_ONLY
        self.family_formula_only_action = self.ACTION_KEEP_TARGET_FORMULA
        self.missing_parameter_action = self.ACTION_ADD_WITH_VALUE

    def action_for_formula_category(self, category):
        if category == 'conflict':
            return self.formula_conflict_action
        if category == 'yaml_only':
            return self.yaml_formula_only_action
        if category == 'family_only':
            return self.family_formula_only_action
        return None


import_alerts = ImportAlertTracker()
import_resolutions = ImportResolutionOptions()

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
    """Return an already-open family document that matches the given DB.Family, if any."""
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

def _collect_project_families(project_doc):
    return list(DB.FilteredElementCollector(project_doc).OfClass(DB.Family))

def _collect_project_family_id_set(project_doc):
    return set([f.Id.IntegerValue for f in _collect_project_families(project_doc)])

def _find_family_by_name_candidates(project_doc, name_candidates):
    cleaned = []
    for name in name_candidates or []:
        if not name:
            continue
        n = name.strip()
        if n and n.lower() not in [x.lower() for x in cleaned]:
            cleaned.append(n)
    if not cleaned:
        return None

    families = _collect_project_families(project_doc)
    # Exact pass
    for cand in cleaned:
        for fam in families:
            try:
                if fam.Name.strip().lower() == cand.lower():
                    return fam
            except Exception:
                continue
    return None

def _resolve_post_load_family(project_doc, pre_family_ids, name_candidates):
    families = _collect_project_families(project_doc)
    new_families = [f for f in families if f.Id.IntegerValue not in (pre_family_ids or set())]
    if len(new_families) == 1:
        return new_families[0]
    if len(new_families) > 1:
        picked = _find_family_by_name_candidates(project_doc, name_candidates)
        if picked:
            return picked
        return new_families[0]
    return _find_family_by_name_candidates(project_doc, name_candidates)

def _get_temp_dir():
    for env_name in ('TEMP', 'TMP'):
        tmp_path = os.environ.get(env_name)
        if tmp_path and os.path.isdir(tmp_path):
            return tmp_path
    return os.path.expanduser('~')

def _make_local_family_copy(source_path):
    temp_dir = _get_temp_dir()
    base_name = os.path.basename(source_path)
    stem, ext = os.path.splitext(base_name)
    local_name = '{}{}'.format(stem, ext or '.rfa')
    local_path = os.path.join(temp_dir, local_name)

    # Prefer original name; if blocked, fall back to a copy suffix.
    if os.path.exists(local_path):
        try:
            os.remove(local_path)
        except Exception:
            alt_name = '{}_copy{}'.format(stem, ext or '.rfa')
            local_path = os.path.join(temp_dir, alt_name)

    shutil.copy2(source_path, local_path)
    return local_path

def _normalize_load_result(load_result):
    """Return tuple(bool_ok, family_obj_or_none, printable_result)."""
    if isinstance(load_result, tuple):
        ok = bool(load_result[0]) if len(load_result) > 0 else False
        fam = load_result[1] if len(load_result) > 1 else None
        return ok, fam, repr(load_result)
    return bool(load_result), None, repr(load_result)

def _probe_file_access(path_to_file):
    try:
        with open(path_to_file, 'rb') as probe:
            probe.read(1)
        return True, ''
    except Exception as ex:
        return False, str(ex)

def _ensure_accessible_rfa_path(path_to_file):
    if not path_to_file:
        return None

    ok, err = _probe_file_access(path_to_file)
    if ok:
        return path_to_file

    retry_value = forms.ask_for_string(
        default=path_to_file,
        prompt='The selected file is not locally accessible.\n'
               'Pick or paste a local .rfa copy path to continue.\n\n'
               'Current error:\n{}'.format(err),
        title='Choose Local RFA Copy'
    )
    if not retry_value:
        return None

    retry_path = _validate_rfa_path(
        os.path.normpath(retry_value.strip().strip('"').strip("'")),
        show_alert=True
    )
    if not retry_path:
        return None

    ok2, err2 = _probe_file_access(retry_path)
    if not ok2:
        forms.alert(
            'Selected file is still not accessible:\n{}\n\n{}'.format(retry_path, err2),
            title='Import Family Types',
            warn_icon=True
        )
        return None

    return retry_path

def _try_load_family(project_doc, path_to_load, loader):
    with revit.Transaction('Load Family From File', doc=project_doc):
        try:
            # Some API versions expose (path, loadOptions), others path-only.
            raw_result = project_doc.LoadFamily(path_to_load, loader)
        except TypeError:
            raw_result = project_doc.LoadFamily(path_to_load)
    return _normalize_load_result(raw_result)

def _load_family_from_file_and_get_ref(project_doc, family_rfa_file, loader, preferred_names=None):
    """Load a family file into the project and return the resulting DB.Family or raise."""
    if not family_rfa_file or not os.path.isfile(family_rfa_file):
        raise Exception('Family file not found: {}'.format(family_rfa_file))

    pre_families = _collect_project_families(project_doc)
    pre_ids = set([f.Id.IntegerValue for f in pre_families])
    attempt_logs = []

    # Attempt 1: direct load from selected path.
    try:
        ok, loaded_family, raw_repr = _try_load_family(project_doc, family_rfa_file, loader)
        attempt_logs.append('direct load: {}'.format(raw_repr))
        if loaded_family:
            return loaded_family
    except Exception as direct_error:
        attempt_logs.append('direct load exception: {}'.format(direct_error))

    # Attempt 2: local temp copy fallback (helps with ACC/connector paths).
    local_copy = None
    try:
        local_copy = _make_local_family_copy(family_rfa_file)
        ok2, loaded_family2, raw_repr2 = _try_load_family(project_doc, local_copy, loader)
        attempt_logs.append('local copy load: {}'.format(raw_repr2))
        if loaded_family2:
            return loaded_family2
    except Exception as copy_error:
        attempt_logs.append('local copy load exception: {}'.format(copy_error))
    finally:
        if local_copy and os.path.exists(local_copy):
            try:
                os.remove(local_copy)
            except Exception:
                pass

    post_families = _collect_project_families(project_doc)
    new_families = [f for f in post_families if f.Id.IntegerValue not in pre_ids]
    if new_families:
        return new_families[0]

    # Fallback: resolve by family name if no new element id was created.
    for family_name in preferred_names or []:
        fam = _find_family_by_name(project_doc, family_name)
        if fam:
            return fam

    raise Exception(
        'LoadFamily failed for file: {}\n{}'.format(
            family_rfa_file,
            '\n'.join(attempt_logs)
        )
    )

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

def _normalize_formula_text(formula_value):
    if formula_value is None:
        return None
    try:
        text_value = str(formula_value).strip()
    except Exception:
        return None
    return text_value or None

def _flag_formula_mismatch(pcfg, fparam):
    yaml_formula = _normalize_formula_text(pcfg.formula)
    try:
        fam_formula = _normalize_formula_text(fparam.Formula)
    except Exception:
        fam_formula = None

    if yaml_formula and fam_formula:
        if yaml_formula != fam_formula:
            import_alerts.add_formula_conflict(pcfg.name, yaml_formula, fam_formula)
        return

    if yaml_formula and not fam_formula:
        import_alerts.add_yaml_formula_only(pcfg.name, yaml_formula)
        return

    if fam_formula and not yaml_formula:
        import_alerts.add_family_formula_only(pcfg.name, fam_formula)

def precheck_parameter_alerts(doc, fconfig):
    param_cfgs = fconfig.get(PARAM_SECTION_NAME, None) or {}
    param_map = _get_param_by_name_map(doc) or {}

    yaml_param_names = set()
    for pname, popts in param_cfgs.items():
        if not (pname and popts):
            continue

        yaml_param_names.add(pname)
        pcfg = get_param_config(doc, pname, popts)
        if not pcfg:
            continue

        fparam = param_map.get(pname)
        if not fparam:
            import_alerts.add_missing_parameter(pname)
            continue

        if bool(fparam.IsInstance) != bool(pcfg.isinst):
            import_alerts.add_instance_mismatch(pname, pcfg.isinst, fparam.IsInstance)

        _flag_formula_mismatch(pcfg, fparam)

    for fname in param_map.keys():
        if fname not in yaml_param_names:
            import_alerts.add_extra_family_parameter(fname)

# ------------------------- Resolution helpers --------------------------------

def _build_discrepancy_summary_text(alerts):
    def _fmt(value):
        return '(none)' if value in (None, '') else str(value)

    lines = ['Discrepancies detected between YAML and the target family:', '']

    if alerts.formula_conflicts:
        lines.append('Formula differences ({}):'.format(len(alerts.formula_conflicts)))
        for pname, payload in sorted(alerts.formula_conflicts.items()):
            lines.append('  - {0}: YAML "{1}" vs Family "{2}"'.format(
                pname, _fmt(payload.get('yaml')), _fmt(payload.get('family'))
            ))
        lines.append('')

    if alerts.yaml_formula_only:
        lines.append('YAML-only formulas ({}):'.format(len(alerts.yaml_formula_only)))
        for pname, formula in sorted(alerts.yaml_formula_only.items()):
            lines.append('  - {0}: YAML formula "{1}"'.format(pname, _fmt(formula)))
        lines.append('')

    if alerts.family_formula_only:
        lines.append('Family-only formulas ({}):'.format(len(alerts.family_formula_only)))
        for pname, formula in sorted(alerts.family_formula_only.items()):
            lines.append('  - {0}: Family formula "{1}"'.format(pname, _fmt(formula)))
        lines.append('')

    if alerts.missing_parameters:
        lines.append('Parameters in YAML but not in family ({}):'.format(len(alerts.missing_parameters)))
        for pname in sorted(alerts.missing_parameters):
            lines.append('  - {0}'.format(pname))
        lines.append('')

    if alerts.extra_family_parameters:
        lines.append('Parameters in family but not in YAML ({}):'.format(len(alerts.extra_family_parameters)))
        for pname in sorted(alerts.extra_family_parameters):
            lines.append('  - {0}'.format(pname))
        lines.append('')

    if alerts.instance_mismatches:
        lines.append('Instance vs Type mismatches ({}):'.format(len(alerts.instance_mismatches)))
        for alert in alerts.instance_mismatches:
            lines.append('  - {param}: YAML {yaml}, Family {family}'.format(
                param=alert['parameter'],
                yaml='Instance' if alert['yaml_isinst'] else 'Type',
                family='Instance' if alert['family_isinst'] else 'Type'
            ))
        lines.append('')

    summary_text = '\n'.join(lines).strip()
    return summary_text or 'No discrepancies detected.'


class DiscrepancyResolutionWindow(forms.WPFWindow):
    def __init__(self, summary_text, alerts, resolutions):
        forms.WPFWindow.__init__(self, RESOLUTION_WINDOW_XAML)
        self._result = None
        self._combo_maps = {}
        self.summary_box.Text = summary_text

        self._init_combo(
            panel=self.conflict_panel,
            combo=self.conflict_combo,
            is_visible=bool(alerts.formula_conflicts),
            option_pairs=[
                (ImportResolutionOptions.ACTION_USE_YAML_FORMULA, 'Use YAML formula'),
                (ImportResolutionOptions.ACTION_KEEP_TARGET_FORMULA, 'Use target formula'),
                (ImportResolutionOptions.ACTION_CLEAR_AND_USE_VALUE, 'Clear target formula and use YAML value')
            ],
            current_value=resolutions.formula_conflict_action
        )

        self._init_combo(
            panel=self.yaml_only_panel,
            combo=self.yaml_only_combo,
            is_visible=bool(alerts.yaml_formula_only),
            option_pairs=[
                (ImportResolutionOptions.ACTION_USE_YAML_FORMULA, 'Use YAML formula'),
                (ImportResolutionOptions.ACTION_USE_YAML_VALUE_ONLY, 'Use YAML value only (no formula)')
            ],
            current_value=resolutions.yaml_formula_only_action
        )

        self._init_combo(
            panel=self.family_only_panel,
            combo=self.family_only_combo,
            is_visible=bool(alerts.family_formula_only),
            option_pairs=[
                (ImportResolutionOptions.ACTION_KEEP_TARGET_FORMULA, 'Use target formula'),
                (ImportResolutionOptions.ACTION_CLEAR_AND_USE_VALUE, 'Clear target formula and use YAML value')
            ],
            current_value=resolutions.family_formula_only_action
        )

        self._init_combo(
            panel=self.missing_panel,
            combo=self.missing_combo,
            is_visible=bool(alerts.missing_parameters),
            option_pairs=[
                (ImportResolutionOptions.ACTION_ADD_WITH_FORMULA, 'Add parameters and formulas'),
                (ImportResolutionOptions.ACTION_ADD_WITH_VALUE, 'Add parameters and values'),
                (ImportResolutionOptions.ACTION_IGNORE_MISSING, 'Ignore parameters (do not add)')
            ],
            current_value=resolutions.missing_parameter_action
        )

    def _init_combo(self, panel, combo, is_visible, option_pairs, current_value):
        if not is_visible:
            panel.Visibility = Visibility.Collapsed
            return
        mapping = {}
        default_label = None
        for value, label in option_pairs:
            combo.Items.Add(label)
            mapping[label] = value
            if value == current_value:
                default_label = label
        if default_label:
            combo.SelectedItem = default_label
        elif combo.Items.Count > 0:
            combo.SelectedIndex = 0
        self._combo_maps[combo.Name] = mapping

    def _get_combo_value(self, combo, fallback):
        mapping = self._combo_maps.get(combo.Name)
        if not mapping:
            return fallback
        selected = combo.SelectedItem
        return mapping.get(selected, fallback)

    def on_ok(self, sender, args):
        self._result = {
            'formula_conflict_action': self._get_combo_value(
                self.conflict_combo, ImportResolutionOptions.ACTION_KEEP_TARGET_FORMULA),
            'yaml_formula_only_action': self._get_combo_value(
                self.yaml_only_combo, ImportResolutionOptions.ACTION_USE_YAML_VALUE_ONLY),
            'family_formula_only_action': self._get_combo_value(
                self.family_only_combo, ImportResolutionOptions.ACTION_KEEP_TARGET_FORMULA),
            'missing_parameter_action': self._get_combo_value(
                self.missing_combo, ImportResolutionOptions.ACTION_ADD_WITH_VALUE)
        }
        self.DialogResult = True
        self.Close()

    def on_cancel(self, sender, args):
        self._result = None
        self.DialogResult = False
        self.Close()

    def get_result(self):
        return self._result

def _set_parameter_formula(fm, fparam, formula_text):
    if formula_text in (None, ''):
        _clear_parameter_formula(fm, fparam)
        return
    try:
        fm.SetFormula(fparam, formula_text)
    except Exception as ex:
        logger.error('Failed to set formula for %s | %s', fparam.Definition.Name, ex)

def _clear_parameter_formula(fm, fparam):
    try:
        fm.SetFormula(fparam, None)
    except Exception:
        try:
            fm.SetFormula(fparam, '')
        except Exception as ex:
            logger.error('Failed to clear formula for %s | %s', fparam.Definition.Name, ex)

def _apply_formula_resolution(pcfg, fparam, fm, was_missing):
    current_formula = _normalize_formula_text(getattr(fparam, 'Formula', None))
    action = None

    if was_missing and import_alerts.is_missing_parameter(pcfg.name):
        missing_action = import_resolutions.missing_parameter_action
        if missing_action == ImportResolutionOptions.ACTION_ADD_WITH_FORMULA and pcfg.formula:
            action = ImportResolutionOptions.ACTION_USE_YAML_FORMULA
        elif missing_action == ImportResolutionOptions.ACTION_ADD_WITH_VALUE and pcfg.formula:
            action = ImportResolutionOptions.ACTION_USE_YAML_VALUE_ONLY
        elif missing_action == ImportResolutionOptions.ACTION_IGNORE_MISSING:
            return bool(current_formula)

    if not action:
        category = import_alerts.get_formula_category(pcfg.name)
        action = import_resolutions.action_for_formula_category(category)

    if action == ImportResolutionOptions.ACTION_USE_YAML_FORMULA:
        if pcfg.formula:
            _set_parameter_formula(fm, fparam, pcfg.formula)
            return True
        return bool(current_formula)

    if action == ImportResolutionOptions.ACTION_KEEP_TARGET_FORMULA:
        return bool(current_formula)

    if action in (ImportResolutionOptions.ACTION_CLEAR_AND_USE_VALUE,
                  ImportResolutionOptions.ACTION_USE_YAML_VALUE_ONLY):
        _clear_parameter_formula(fm, fparam)
        return False

    # default: respect whatever the family currently has
    return bool(current_formula)

def resolve_discrepancy_actions():
    import_resolutions.reset()

    if not import_alerts.has_actionable_alerts():
        return True

    summary_text = _build_discrepancy_summary_text(import_alerts)

    window = DiscrepancyResolutionWindow(summary_text, import_alerts, import_resolutions)
    dialog_result = window.ShowDialog()
    if not dialog_result:
        return False

    selections = window.get_result()
    if not selections:
        return False

    if import_alerts.formula_conflicts:
        import_resolutions.formula_conflict_action = selections.get(
            'formula_conflict_action', import_resolutions.formula_conflict_action)
    if import_alerts.yaml_formula_only:
        import_resolutions.yaml_formula_only_action = selections.get(
            'yaml_formula_only_action', import_resolutions.yaml_formula_only_action)
    if import_alerts.family_formula_only:
        import_resolutions.family_formula_only_action = selections.get(
            'family_formula_only_action', import_resolutions.family_formula_only_action)
    if import_alerts.missing_parameters:
        import_resolutions.missing_parameter_action = selections.get(
            'missing_parameter_action', import_resolutions.missing_parameter_action)

    return True

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

        param_was_missing = import_alerts.is_missing_parameter(pname)
        if param_was_missing and import_resolutions.missing_parameter_action == ImportResolutionOptions.ACTION_IGNORE_MISSING:
            logger.debug('Skipping YAML-only parameter per user choice: %s', pname)
            continue

        # Skip instance parameters entirely if toggle enabled
        if SKIP_INSTANCE_PARAMS and pcfg.isinst:
            logger.debug('Skipping instance parameter by mode: %s', pname)
            continue

        fparam = ensure_param(doc, fm, pcfg, pname)
        if not fparam:
            continue

        _apply_formula_resolution(pcfg, fparam, fm, param_was_missing)

        if pcfg.default is not None or pcfg.isreport:
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

def _normalize_type_name(value):
    return (value or '').strip().lower()

def _get_family_type_by_name(fm, type_name):
    target = _normalize_type_name(type_name)
    for ftype in fm.Types:
        try:
            current = _normalize_type_name(ftype.Name)
        except Exception:
            current = _normalize_type_name(str(ftype.Name))
        if current == target:
            return ftype
    return None

def purge_unselected_types(doc, fconfig):
    """Delete any family type not included in the selected import config."""
    fm = doc.FamilyManager
    keep_cfg = fconfig.get(TYPES_SECTION_NAME, {}) or {}
    keep_names = set([_normalize_type_name(k) for k in keep_cfg.keys() if _normalize_type_name(k)])
    if not keep_names:
        return 0, []

    # Snapshot current types before deleting.
    existing_types = []
    for ftype in fm.Types:
        try:
            tname = ftype.Name
        except Exception:
            tname = str(ftype.Name)
        existing_types.append((ftype, tname))

    delete_targets = []
    for ftype, tname in existing_types:
        if _normalize_type_name(tname) in keep_names:
            continue
        delete_targets.append((ftype, tname))

    if not delete_targets:
        return 0, []

    deleted_count = 0
    remaining_targets = list(delete_targets)
    max_passes = 5
    for _ in range(max_passes):
        if not remaining_targets:
            break
        progress = 0
        next_remaining = []

        for ftype, tname in remaining_targets:
            # Revit does not allow deleting the final remaining type.
            current_count = len([t for t in fm.Types])
            if current_count <= 1:
                next_remaining.append((ftype, tname))
                continue
            try:
                # Resolve by name in case the old ftype handle is stale.
                current = _get_family_type_by_name(fm, tname) or ftype
                fm.CurrentType = current
                fm.DeleteCurrentType()
                deleted_count += 1
                progress += 1
            except Exception as ex:
                logger.warning('Could not delete unselected type "%s" | %s', tname, ex)
                next_remaining.append((ftype, tname))

        remaining_targets = next_remaining
        if progress == 0:
            break

    leftover = []
    for ftype in fm.Types:
        try:
            tname = ftype.Name
        except Exception:
            tname = str(ftype.Name)
        if _normalize_type_name(tname) not in keep_names:
            leftover.append(tname)

    return deleted_count, sorted(leftover, key=lambda n: (n or '').lower())

def _get_family_symbol_name(symbol):
    try:
        name_param = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if name_param:
            pname = name_param.AsString()
            if pname:
                return pname
    except Exception:
        pass
    try:
        return symbol.Name
    except Exception:
        return ''

def _get_project_family_symbols(project_doc, family_obj):
    symbols = []
    try:
        symbol_ids = list(family_obj.GetFamilySymbolIds())
    except Exception:
        symbol_ids = []
    for sid in symbol_ids:
        sym = project_doc.GetElement(sid)
        if sym:
            symbols.append(sym)
    return symbols

def _pick_fallback_keep_symbol(symbols, keep_names, exclude_id=None):
    exclude_int = exclude_id.IntegerValue if exclude_id else None
    for sym in symbols:
        try:
            sid_int = sym.Id.IntegerValue
        except Exception:
            sid_int = None
        if exclude_int is not None and sid_int == exclude_int:
            continue
        if _normalize_type_name(_get_family_symbol_name(sym)) in keep_names:
            return sym
    for sym in symbols:
        try:
            sid_int = sym.Id.IntegerValue
        except Exception:
            sid_int = None
        if exclude_int is not None and sid_int == exclude_int:
            continue
        return sym
    return None

def _reassign_instances_from_symbol(project_doc, from_symbol, to_symbol):
    if not from_symbol or not to_symbol:
        return 0
    moved = 0
    for inst in DB.FilteredElementCollector(project_doc).OfClass(DB.FamilyInstance):
        try:
            if not inst.Symbol:
                continue
            if inst.Symbol.Id != from_symbol.Id:
                continue
            inst.Symbol = to_symbol
            moved += 1
        except Exception:
            continue
    return moved

def purge_unselected_project_types(project_doc, family_obj, fconfig):
    """Delete unselected types from the loaded family inside the project document."""
    if not family_obj:
        return 0, []

    keep_cfg = fconfig.get(TYPES_SECTION_NAME, {}) or {}
    keep_names = set([_normalize_type_name(k) for k in keep_cfg.keys() if _normalize_type_name(k)])
    if not keep_names:
        return 0, []

    symbol_ids = [sym.Id for sym in _get_project_family_symbols(project_doc, family_obj)]
    if not symbol_ids:
        return 0, []

    delete_ids = []
    for sid in symbol_ids:
        sym = project_doc.GetElement(sid)
        if not sym:
            continue
        sname = _normalize_type_name(_get_family_symbol_name(sym))
        if sname in keep_names:
            continue
        delete_ids.append(sid)

    if not delete_ids:
        return 0, []

    deleted = 0
    remaining_ids = [sid for sid in delete_ids]
    max_passes = 5
    with revit.Transaction('Purge Unselected Project Family Types', doc=project_doc):
        for _ in range(max_passes):
            if not remaining_ids:
                break
            progress = 0
            next_remaining = []

            for sid in remaining_ids:
                current_symbols = _get_project_family_symbols(project_doc, family_obj)
                remaining_count = len(current_symbols)

                # Keep at least one type to avoid invalid family state.
                if remaining_count <= 1:
                    next_remaining.append(sid)
                    continue

                try:
                    delete_symbol = project_doc.GetElement(sid)
                    if not delete_symbol:
                        progress += 1
                        continue

                    keep_symbol = _pick_fallback_keep_symbol(
                        current_symbols,
                        keep_names,
                        exclude_id=sid
                    )
                    if keep_symbol and keep_symbol.Id != sid:
                        _reassign_instances_from_symbol(project_doc, delete_symbol, keep_symbol)

                    project_doc.Delete(sid)
                    deleted += 1
                    progress += 1
                except Exception as ex:
                    sym = project_doc.GetElement(sid)
                    sname = _get_family_symbol_name(sym) if sym else str(sid.IntegerValue)
                    logger.warning('Could not delete project type "%s" | %s', sname, ex)
                    next_remaining.append(sid)

            remaining_ids = next_remaining
            if progress == 0:
                break

    leftover_unselected = []
    for sym in _get_project_family_symbols(project_doc, family_obj):
        sname = _get_family_symbol_name(sym)
        if _normalize_type_name(sname) not in keep_names:
            leftover_unselected.append(sname)

    return deleted, sorted(leftover_unselected, key=lambda n: (n or '').lower())

# ------------------------- YAML & Type Picker --------------------------------

def _clean_type_list(type_values):
    if not isinstance(type_values, list):
        return []
    cleaned = []
    for tname in type_values:
        if not isinstance(tname, basestring):
            continue
        name = tname.strip()
        if name:
            cleaned.append(name)
    return sorted(set(cleaned), key=lambda x: x.lower())

def get_map_file():
    init_dir = FAMILY_CONFIGS_ROOT if os.path.isdir(FAMILY_CONFIGS_ROOT) else None
    map_file = forms.pick_file(
        file_ext='yaml',
        title='Choose map.yaml',
        init_dir=init_dir
    )
    if not map_file:
        return None

    if os.path.basename(map_file).lower() != MAP_FILE_NAME:
        forms.alert(
            'Please select "{}".'.format(MAP_FILE_NAME),
            title='Import Family Types',
            warn_icon=True
        )
        return None
    return map_file

def ask_include_default_types_option():
    option_label = 'Include all existing/default types (skip purge)'
    selected = forms.SelectFromList.show(
        [option_label],
        title='Type Retention Option',
        button_name='Continue',
        multiselect=True,
        message='Check this option to keep all existing/default family types.\n'
                'Leave unchecked to load only the selected types.'
    )
    if selected is None:
        return None
    # pyRevit may return wrapped list items rather than raw strings.
    # Since there is a single checkbox option, any selection means "include defaults".
    try:
        return len(selected) > 0
    except Exception:
        return bool(selected)

def _validate_rfa_path(candidate_path, show_alert=True):
    if not candidate_path:
        return None
    normalized = os.path.normpath(candidate_path)
    if not normalized.lower().endswith('.rfa'):
        if show_alert:
            forms.alert(
                'Please select a .rfa file.',
                title='Import Family Types',
                warn_icon=True
            )
        return None
    if not os.path.isfile(normalized):
        if show_alert:
            forms.alert(
                'File not found:\n{}'.format(normalized),
                title='Import Family Types',
                warn_icon=True
            )
        return None
    return normalized

def get_family_file():
    init_dir = RFA_PICKER_DEFAULT_DIR if os.path.isdir(RFA_PICKER_DEFAULT_DIR) else None
    family_file = forms.pick_file(
        file_ext='rfa',
        title='Choose RFA file',
        init_dir=init_dir
    )
    valid = _validate_rfa_path(family_file, show_alert=False)
    if not valid:
        forms.alert('No RFA selected, closing ImportFamilyTypes', title='Import Family Types')
        return None
    return valid

def _family_name_from_rfa_path(family_rfa_file):
    return os.path.splitext(os.path.basename(family_rfa_file))[0].strip()

def load_map_configs(map_file):
    map_configs = yaml.load_as_dict(map_file)
    if not isinstance(map_configs, dict):
        forms.alert(
            '"{}" is not a valid map file.'.format(MAP_FILE_NAME),
            title='Import Family Types',
            warn_icon=True
        )
        return None
    return map_configs

def _find_yaml_candidates_for_family(map_configs, family_name):
    candidates = []
    family_key = (family_name or '').strip().lower()
    if not family_key:
        return candidates

    for yaml_key, type_list in map_configs.items():
        if not isinstance(yaml_key, basestring):
            continue
        yaml_key_clean = yaml_key.strip()
        if not yaml_key_clean:
            continue

        yaml_stem = os.path.splitext(os.path.basename(yaml_key_clean))[0].strip()
        yaml_stem_lower = yaml_stem.lower()
        score = None

        if yaml_stem_lower == family_key:
            score = 0
        elif yaml_stem_lower.startswith(family_key + '_') \
                or yaml_stem_lower.startswith(family_key + '-') \
                or yaml_stem_lower.startswith(family_key + ' '):
            score = 1
        elif yaml_stem_lower.startswith(family_key):
            score = 2

        if score is None:
            continue

        candidates.append({
            'map_key': yaml_key_clean,
            'score': score,
            'types': _clean_type_list(type_list),
            'stem': yaml_stem,
        })

    candidates = sorted(candidates, key=lambda c: (c['score'], len(c['stem']), c['stem'].lower()))
    return candidates

def _resolve_yaml_path_from_map(map_file, yaml_key):
    map_root = os.path.dirname(map_file)
    source_root = os.path.join(map_root, 'Source')
    key_clean = (yaml_key or '').strip()
    if not key_clean:
        return None

    key_basename = os.path.basename(key_clean)
    key_basename_lower = key_basename.lower()
    key_stem_lower = os.path.splitext(key_basename_lower)[0]

    # Preferred resolution order:
    # 1) sibling "Source" folder (expected layout)
    # 2) map root folder
    candidate_roots = []
    for root in [source_root, map_root]:
        if root and os.path.isdir(root) and root not in candidate_roots:
            candidate_roots.append(root)

    # If the key is a relative or nested path, try direct resolution first per root.
    for root in candidate_roots:
        direct_candidate = os.path.normpath(os.path.join(root, key_clean))
        if os.path.exists(direct_candidate):
            return direct_candidate
        by_name_candidate = os.path.normpath(os.path.join(root, key_basename))
        if os.path.exists(by_name_candidate):
            return by_name_candidate

    matches = []
    for root in candidate_roots:
        for walk_root, _, files in os.walk(root):
            for filename in files:
                file_lower = filename.lower()
                if file_lower == key_basename_lower:
                    matches.append(os.path.join(walk_root, filename))

        if not matches and key_basename_lower.endswith('.yaml'):
            for walk_root, _, files in os.walk(root):
                for filename in files:
                    file_lower = filename.lower()
                    if file_lower.endswith('.yaml') \
                            and os.path.splitext(file_lower)[0] == key_stem_lower:
                        matches.append(os.path.join(walk_root, filename))

    if not matches:
        return None

    matches = sorted(set(matches), key=lambda p: (len(p), p.lower()))
    if len(matches) == 1:
        return matches[0]
    logger.warning(
        'Multiple YAML files found for key "%s". Using: %s',
        key_basename,
        matches[0]
    )
    return matches[0]

def pick_types_from_map_candidates(candidates):
    type_to_map_keys = {}
    type_display = {}

    for candidate in candidates:
        map_key = candidate.get('map_key')
        for type_name in candidate.get('types', []):
            normalized = (type_name or '').strip().lower()
            if not normalized:
                continue
            if normalized not in type_to_map_keys:
                type_to_map_keys[normalized] = []
                type_display[normalized] = type_name.strip()
            type_to_map_keys[normalized].append(map_key)

    all_type_labels = sorted(type_display.values(), key=lambda n: n.lower())
    if not all_type_labels:
        forms.alert(
            'No types were found in map entries for this family.',
            title='Type Picker',
            warn_icon=True
        )
        return None, None

    selected = forms.SelectFromList.show(
        all_type_labels,
        title='Select Family Types to Import',
        button_name='Load Selected Types',
        multiselect=True
    )
    if selected is None:
        return None, None

    if len(selected) > 50:
        res = forms.alert(
            'You selected {} types.\nLarge imports can be very slow.\n\n'
            'Continue anyway?'.format(len(selected)),
            title='Confirm Large Import',
            warn_icon=True,
            yes=True, no=True
        )
        if not res:
            return None, None

    selected_type_to_maps = {}
    selected_labels = []
    for label in selected:
        normalized = (label or '').strip().lower()
        if not normalized or normalized not in type_to_map_keys:
            continue
        selected_labels.append(type_display[normalized])
        selected_type_to_maps[type_display[normalized]] = list(type_to_map_keys[normalized])

    return selected_labels, selected_type_to_maps

def _merge_param_sections(base_params, incoming_params):
    if not isinstance(incoming_params, dict):
        return
    for pname, popts in incoming_params.items():
        if pname not in base_params:
            base_params[pname] = popts
            continue
        if base_params[pname] != popts:
            logger.warning(
                'Parameter definition mismatch for "%s" between YAML sources; keeping first definition.',
                pname
            )

def _lookup_type_payload(types_dict, type_name):
    if not isinstance(types_dict, dict):
        return None, None
    target = (type_name or '').strip().lower()
    for key, payload in types_dict.items():
        if (key or '').strip().lower() == target:
            return key, payload
    return None, None

def build_family_config_from_map_selection(map_file, family_name, selected_type_to_maps):
    required_map_keys = []
    for _, map_keys in selected_type_to_maps.items():
        for map_key in map_keys:
            if map_key not in required_map_keys:
                required_map_keys.append(map_key)

    yaml_sources = []
    loaded_cfg_by_map_key = {}
    for map_key in required_map_keys:
        yaml_path = _resolve_yaml_path_from_map(map_file, map_key)
        if not yaml_path:
            forms.alert(
                'Could not locate YAML file "{}" near {}.'.format(map_key, MAP_FILE_NAME),
                title='Import Family Types',
                warn_icon=True
            )
            return None, []
        cfg = load_configs(yaml_path)
        if not isinstance(cfg, dict):
            forms.alert(
                'YAML is invalid:\n{}'.format(yaml_path),
                title='Import Family Types',
                warn_icon=True
            )
            return None, []
        loaded_cfg_by_map_key[map_key] = cfg
        yaml_sources.append(yaml_path)

    combined = {
        'family': family_name,
        PARAM_SECTION_NAME: {},
        TYPES_SECTION_NAME: {},
        SHAREDPARAM_DEF: '',
    }
    unresolved_types = []

    for map_key in required_map_keys:
        cfg = loaded_cfg_by_map_key[map_key]
        _merge_param_sections(
            combined[PARAM_SECTION_NAME],
            cfg.get(PARAM_SECTION_NAME, {})
        )
        shared_defs = cfg.get(SHAREDPARAM_DEF)
        if shared_defs and not combined[SHAREDPARAM_DEF]:
            combined[SHAREDPARAM_DEF] = shared_defs

    for type_name, map_keys in selected_type_to_maps.items():
        found = False
        for map_key in map_keys:
            cfg = loaded_cfg_by_map_key.get(map_key, {})
            type_key, type_payload = _lookup_type_payload(
                cfg.get(TYPES_SECTION_NAME, {}),
                type_name
            )
            if type_key is None:
                continue
            combined[TYPES_SECTION_NAME][type_key] = type_payload
            found = True
            break
        if not found:
            unresolved_types.append(type_name)

    if unresolved_types:
        forms.alert(
            'Some selected types could not be found in source YAMLs:\n{}'.format(
                '\n'.join(sorted(unresolved_types, key=lambda n: n.lower()))
            ),
            title='Import Family Types',
            warn_icon=True
        )

    if not combined[TYPES_SECTION_NAME]:
        forms.alert(
            'No selected types could be resolved from map.yaml.',
            title='Import Family Types',
            warn_icon=True
        )
        return None, []

    return combined, yaml_sources

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
    map_file = get_map_file()
    if not map_file:
        raise SystemExit

    family_rfa_file = get_family_file()
    if not family_rfa_file:
        raise SystemExit

    project_doc = revit.doc
    is_family_doc = project_doc.IsFamilyDocument

    selected_family_name = _family_name_from_rfa_path(family_rfa_file)
    if not selected_family_name:
        forms.alert('Could not determine family name from selected RFA file.', warn_icon=True)
        raise SystemExit

    # If this family is not already loaded in the project, verify source file access
    # up-front so users don't spend time picking types before a guaranteed load failure.
    if (not is_family_doc) and (not _find_family_by_name(project_doc, selected_family_name)):
        verified_path = _ensure_accessible_rfa_path(family_rfa_file)
        if not verified_path:
            forms.alert('Import cancelled: accessible family file is required.', title='Import Family Types')
            raise SystemExit
        family_rfa_file = verified_path

    map_configs = load_map_configs(map_file)
    if map_configs is None:
        raise SystemExit

    candidate_entries = _find_yaml_candidates_for_family(map_configs, selected_family_name)
    if not candidate_entries:
        forms.alert(
            'No map entries found for family "{}".'.format(selected_family_name),
            title='Import Family Types',
            warn_icon=True
        )
        raise SystemExit

    _, selected_type_to_maps = pick_types_from_map_candidates(candidate_entries)
    if selected_type_to_maps is None:
        forms.alert('Import cancelled.', title='Import Family Types')
        raise SystemExit

    include_default_types = ask_include_default_types_option()
    if include_default_types is None:
        forms.alert('Import cancelled.', title='Import Family Types')
        raise SystemExit

    family_configs, source_yaml_files = build_family_config_from_map_selection(
        map_file,
        selected_family_name,
        selected_type_to_maps
    )
    if not family_configs:
        raise SystemExit

    family_cfg_file = source_yaml_files[0] if source_yaml_files else map_file
    if not isinstance(family_configs, dict):
        forms.alert('Resolved family configuration is invalid.', title='Import Family Types', warn_icon=True)
        raise SystemExit

    import_alerts.reset()
    import_resolutions.reset()
    existing_sharedparam_file = HOST_APP.app.SharedParametersFilename

    fam_doc = None
    target_family = None
    temp_open_copy_path = None
    loader = _AlwaysOverwriteLoader(overwrite_params=True)
    should_close_fam_doc = False
    was_existing_project_family = False

    try:
        yaml_family_name = _get_yaml_family_name(family_configs)
        if yaml_family_name and selected_family_name \
                and yaml_family_name.strip().lower() != selected_family_name.strip().lower():
            logger.warning(
                "Selected RFA family '%s' does not match YAML family '%s'. Continuing with selected RFA family.",
                selected_family_name,
                yaml_family_name
            )

        if is_family_doc:
            # Keep compatibility with family environment.
            fam_doc = project_doc
            try:
                current_name = fam_doc.FamilyManager.OwnerFamily.Name
            except Exception:
                current_name = (fam_doc.Title[:-4] if fam_doc.Title.lower().endswith('.rfa') else fam_doc.Title)
            if selected_family_name and current_name.strip().lower() != selected_family_name.strip().lower():
                logger.warning(
                    "Selected RFA family '%s' does not match open family '%s'. Proceeding with open family.",
                    selected_family_name, current_name
                )
        else:
            target_family = _find_family_by_name(project_doc, selected_family_name)
            if target_family:
                was_existing_project_family = True
                if include_default_types:
                    # Pull in all source-RFA default types first, then apply map-selected types.
                    try:
                        target_family = _load_family_from_file_and_get_ref(
                            project_doc,
                            family_rfa_file,
                            loader,
                            preferred_names=[selected_family_name, yaml_family_name]
                        )
                    except Exception as load_error:
                        forms.alert(
                            'Could not load default types from selected RFA:\n{}\n\n{}'.format(
                                family_rfa_file, load_error
                            ),
                            title='Import Family Types',
                            warn_icon=True
                        )
                        raise SystemExit

                    if not target_family:
                        forms.alert(
                            'Could not resolve family after loading selected RFA defaults.',
                            title='Import Family Types',
                            warn_icon=True
                        )
                        raise SystemExit

                fam_doc = _get_open_family_doc_for(target_family)
                if fam_doc:
                    should_close_fam_doc = False
                else:
                    existing_doc_sigs = _get_open_doc_signatures()
                    fam_doc = project_doc.EditFamily(target_family)
                    should_close_fam_doc = _doc_signature(fam_doc) not in existing_doc_sigs
            else:
                if include_default_types:
                    # Keep source/default types by editing the loaded family directly.
                    try:
                        target_family = _load_family_from_file_and_get_ref(
                            project_doc,
                            family_rfa_file,
                            loader,
                            preferred_names=[selected_family_name, yaml_family_name]
                        )
                    except Exception as load_error:
                        forms.alert(
                            'Could not load family file:\n{}\n\n{}'.format(
                                family_rfa_file, load_error
                            ),
                            title='Import Family Types',
                            warn_icon=True
                        )
                        raise SystemExit

                    if not target_family:
                        forms.alert(
                            'Family file was not loaded into project:\n{}'.format(family_rfa_file),
                            title='Import Family Types',
                            warn_icon=True
                        )
                        raise SystemExit

                    fam_doc = _get_open_family_doc_for(target_family)
                    if fam_doc:
                        should_close_fam_doc = False
                    else:
                        existing_doc_sigs = _get_open_doc_signatures()
                        fam_doc = project_doc.EditFamily(target_family)
                        should_close_fam_doc = _doc_signature(fam_doc) not in existing_doc_sigs
                else:
                    # Purge path: edit a temporary local copy, then load to project.
                    try:
                        temp_open_copy_path = _make_local_family_copy(family_rfa_file)
                        fam_doc = HOST_APP.app.OpenDocumentFile(temp_open_copy_path)
                        should_close_fam_doc = True
                    except Exception as open_error:
                        forms.alert(
                            'Could not open family file for editing:\n{}\n\n{}'.format(
                                family_rfa_file, open_error
                            ),
                            title='Import Family Types',
                            warn_icon=True
                        )
                        raise SystemExit

        if SHAREDPARAM_DEF in family_configs:
            sharedparam_file = recover_sharedparam_defs(family_cfg_file, family_configs[SHAREDPARAM_DEF])
            HOST_APP.app.SharedParametersFilename = sharedparam_file

        precheck_parameter_alerts(fam_doc, family_configs)

        if import_alerts.has_actionable_alerts():
            if not resolve_discrepancy_actions():
                logger.warning('Import cancelled before any changes were made.')
                raise SystemExit
        elif import_alerts.has_alerts():
            import_alerts.report()

        with revit.Transaction('Import Params/Types from Config', doc=fam_doc):
            fam_mgr = fam_doc.FamilyManager
            ctype = fam_mgr.CurrentType or fam_mgr.NewType(TEMP_TYPENAME)
            ctype_name = None
            try:
                ctype_name = ctype.Name if ctype else None
            except Exception:
                ctype_name = None

            ensure_params(fam_doc, family_configs)
            ensure_types(fam_doc, family_configs)

            should_purge_unselected = (not include_default_types) and (not was_existing_project_family)
            if (not is_family_doc) and should_purge_unselected:
                purged, leftover_family = purge_unselected_types(fam_doc, family_configs)
                if purged:
                    logger.debug('Purged %s unselected types from loaded family doc.', purged)
                if leftover_family:
                    logger.warning(
                        'Family doc still has unselected types after purge: %s',
                        ', '.join(leftover_family)
                    )
                    forms.alert(
                        'Some unselected types remained in the family editor doc:\n{}\n\n'
                        'These will also be targeted in project-level purge.'
                        .format('\n'.join(leftover_family)),
                        title='Import Family Types',
                        warn_icon=True
                    )

            if ctype_name and ctype_name != TEMP_TYPENAME:
                restored = _get_family_type_by_name(fam_mgr, ctype_name)
                if restored:
                    fam_mgr.CurrentType = restored

        if not is_family_doc:
            loaded_name_hint = None
            try:
                loaded_name_hint = fam_doc.FamilyManager.OwnerFamily.Name
            except Exception:
                loaded_name_hint = None

            pre_load_family_ids = _collect_project_family_id_set(project_doc)
            fam_doc.LoadFamily(project_doc, loader)

            should_purge_unselected = (not include_default_types) and (not was_existing_project_family)
            if should_purge_unselected:
                # Resolve the just-loaded family robustly (new id, selected name, yaml name, or temp name).
                name_candidates = [selected_family_name, yaml_family_name, loaded_name_hint]
                if temp_open_copy_path:
                    try:
                        name_candidates.append(
                            os.path.splitext(os.path.basename(temp_open_copy_path))[0]
                        )
                    except Exception:
                        pass
                post_family = _resolve_post_load_family(
                    project_doc,
                    pre_load_family_ids,
                    name_candidates
                )

                if post_family:
                    purged_project, leftover_unselected = purge_unselected_project_types(
                        project_doc, post_family, family_configs
                    )
                    if purged_project:
                        logger.debug('Purged %s unselected project family types.', purged_project)
                    if leftover_unselected:
                        logger.warning(
                            'Some unselected project types could not be purged: %s',
                            ', '.join(leftover_unselected)
                        )
                        forms.alert(
                            'Some unselected types could not be removed (likely because Revit locked them).\n\n{}'
                            .format('\n'.join(leftover_unselected)),
                            title='Import Family Types',
                            warn_icon=True
                        )
                else:
                    forms.alert(
                        'Family loaded but post-load family reference could not be resolved.\n'
                        'Project-level type purge was skipped.',
                        title='Import Family Types',
                        warn_icon=True
                    )

            if was_existing_project_family:
                forms.alert('Family updated in project.', title='Import Family Types')
            else:
                forms.alert('Family loaded into project with selected types.', title='Import Family Types')

    except Exception as import_error:
        logger.error('Import failed: %s', import_error)
        raise
    finally:
        HOST_APP.app.SharedParametersFilename = existing_sharedparam_file
        if fam_doc and should_close_fam_doc:
            try:
                fam_doc.Close(False)
            except Exception:
                pass
        if temp_open_copy_path and os.path.exists(temp_open_copy_path):
            try:
                os.remove(temp_open_copy_path)
            except Exception:
                pass
