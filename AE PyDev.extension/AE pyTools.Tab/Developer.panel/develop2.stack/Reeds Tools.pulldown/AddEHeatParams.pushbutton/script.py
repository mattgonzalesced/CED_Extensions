# -*- coding: utf-8 -*-
# Revit 2025+ | pyRevit IronPython | Active Family OR Multi-Family (.rfa) | No Save
# Extras:
# - Opens families without UI noise (ModelPath + OpenOptions)
# - Swallows warning failures during transactions
# - Temporarily auto-dismisses TaskDialogs like "pyRevit – Document Opened: …"
# - Tracks families missing "Power Connection Required"
# - Applies different formulas depending on "Power Connection Required" value
# - Reordering is a dedicated, separate transaction (after add/update)

import clr, os, sys, re
from System.Reflection import BindingFlags
import System
from System import Type, Activator, Array, Guid
from System.Runtime.InteropServices import Marshal

try:
    _SCRIPT_DIR = os.path.dirname(__file__)
except:
    _SCRIPT_DIR = os.getcwd()
if _SCRIPT_DIR not in sys.path:
    sys.path.insert(0, _SCRIPT_DIR)

#from helper_power_balance import run_helper_delete_power_balance_and_save

AUTO_SAVE_NONACTIVE = True   # save families opened by the script (recommended)
AUTO_SAVE_ACTIVE   = False   # set True if you also want to auto-save the active family

def _save_family_doc(doc, path_hint=""):
    """Save the family doc to its current PathName if there are changes."""
    try:
        if doc.IsModified:
            doc.Save()
            try:
                p = doc.PathName or path_hint or "(unknown path)"
            except:
                p = path_hint or "(unknown path)"
            log("[INFO] Saved: {}".format(p))
            return True
        else:
            log("[INFO] No changes; skipped save.")
            return True
    except Exception as e:
        log("[ERROR] Save failed{}: {}".format(
            " for {}".format(path_hint) if path_hint else "", e))
        return False

# Revit API (DB)
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import (
    Transaction, SpecTypeId, GroupTypeId, ExternalDefinition, ForgeTypeId, LabelUtils,
    OpenOptions, ModelPathUtils, IFailuresPreprocessor, FailureSeverity, FailureProcessingResult
)

# Revit API (UI)
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import TaskDialogResult
from Autodesk.Revit.UI.Events import DialogBoxShowingEventArgs

# RevitServices
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager

# WinForms
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import (Form, Label, TextBox, Button, ComboBox, ComboBoxStyle,
                                  OpenFileDialog, Control, DialogResult)
from System.Drawing import Point, Size
from System import Array as DotNetArray

uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

# ignore any leading CC-SSS style prefixes when matching
_prefix_re_local = re.compile(r'^\d{2}-\d{3}(?:-\d+)?\s+')

# ------------------------------ Config / constants ------------------------------
REQ_HEADERS = ["Parameter Name", "Type of Parameter", "Group Under", "Instance/Type", "Formula"]
HDR_GUID, HDR_FAMSH = "GUID", "Family/Shared"
HDR_CIRCUIT, HDR_SORT = "Circuit Group", "Sort Order"

# Group mapping
_GROUP_SIMPLE = {
    "data": GroupTypeId.Data,
    "text": GroupTypeId.Text,
    "constraints": GroupTypeId.Constraints,
    "geometry": GroupTypeId.Geometry,
    "graphics": GroupTypeId.Graphics,
    "identitydata": GroupTypeId.IdentityData,
    "materials": GroupTypeId.Materials,
    "construction": GroupTypeId.Construction,
    "electrical": GroupTypeId.Electrical,
    "electricalloads": GroupTypeId.ElectricalLoads if hasattr(GroupTypeId, "ElectricalLoads") else GroupTypeId.Electrical,
    "plumbing": GroupTypeId.Plumbing if hasattr(GroupTypeId, "Plumbing") else GroupTypeId.Data,
    "mechanical": GroupTypeId.Mechanical if hasattr(GroupTypeId, "Mechanical") else GroupTypeId.Data,
    "mechanicalloads": GroupTypeId.MechanicalLoads if hasattr(GroupTypeId, "MechanicalLoads")
                        else (GroupTypeId.Mechanical if hasattr(GroupTypeId, "Mechanical") else GroupTypeId.Data),
    "energyanalysis": GroupTypeId.EnergyAnalysis if hasattr(GroupTypeId, "EnergyAnalysis") else GroupTypeId.Data,
    "dimensions": GroupTypeId.Geometry,
    "other": GroupTypeId.Other if hasattr(GroupTypeId, "Other") else GroupTypeId.Data,
}

TYPE_ALIAS = {
    "yes/no":"YesNo","yesno":"YesNo","boolean":"YesNo",
    "text":"Text","string":"Text",
    "length":"Length","area":"Area","volume":"Volume",
    "integer":"Integer","int":"Integer",
    "number":"Number","double":"Number","float":"Number",
    "material":"Material","angle":"Angle","slope":"Slope","currency":"Currency",
    "url":"Url","hyperlink":"Url"
}

def log(msg): print(msg)

def _norm_key(s):
    if not s: return ""
    s = s.strip().lower()
    return "".join(ch for ch in s if ch.isalnum())

def group_type_id(label):
    key = _norm_key(label)
    return _GROUP_SIMPLE.get(key, GroupTypeId.Data)

def resolve_spec(s):
    if not s: return SpecTypeId.String.Text
    key = TYPE_ALIAS.get((s or "").strip().lower(), (s or "").strip())
    try:
        if "." in key:
            obj = SpecTypeId
            for part in key.split("."):
                obj = getattr(obj, part)
            return obj
        return getattr(SpecTypeId, key)
    except:
        return SpecTypeId.String.Text

# ------------------------------ Global tracking ------------------------------
POWER_CONN_NOT_PARAM = []  # families missing "Power Connection Required"

def has_param(fam_doc, name):
    fm = fam_doc.FamilyManager
    for fp in fm.GetParameters():
        try:
            if fp.Definition and fp.Definition.Name == name:
                return True
        except:
            pass
    return False

def get_param(fm, name):
    for fp in fm.GetParameters():
        try:
            if fp.Definition and fp.Definition.Name == name:
                return fp
        except:
            pass
    return None

def _dump_power_conn_missing():
    log("\n===== Missing 'Power Connection Required' =====")
    if POWER_CONN_NOT_PARAM:
        for p in POWER_CONN_NOT_PARAM:
            log(" - {}".format(p))
    else:
        log("All processed families have 'Power Connection Required'.")

# ------------------------------ UI ------------------------------
class Runner(Form):
    def __init__(self):
        self.Text = "Add Parameters (Active or Selected Families)"
        self.Size = Size(900, 260)

        # Excel
        self.lblX = Label(Text="Excel"); self.lblX.Location = Point(10, 20)
        self.txtX = TextBox(); self.txtX.Location = Point(140, 18); self.txtX.Size = Size(640, 22)
        self.btnX = Button(Text="Browse"); self.btnX.Location = Point(790, 16); self.btnX.Click += self.pick_xlsx

        # Shared Params
        self.lblS = Label(Text="Shared Params"); self.lblS.Location = Point(10, 55)
        self.txtS = TextBox(); self.txtS.Location = Point(140, 53); self.txtS.Size = Size(640, 22)
        self.btnS = Button(Text="Browse"); self.btnS.Location = Point(790, 51); self.btnS.Click += self.pick_sp

        # Circuits
        self.lblC = Label(Text="Circuits (1-5):"); self.lblC.Location = Point(10, 90)
        self.cboC = ComboBox(); self.cboC.Location = Point(140, 88); self.cboC.DropDownStyle = ComboBoxStyle.DropDownList
        for i in range(1,6): self.cboC.Items.Add(str(i))
        self.cboC.SelectedIndex = 0

        # Multi-family
        self.lblF = Label(Text="Families (.rfa)"); self.lblF.Location = Point(10, 130)
        self.txtF = TextBox(); self.txtF.Location = Point(140, 128); self.txtF.Size = Size(640, 22); self.txtF.ReadOnly = True
        self.btnF = Button(Text="Browse"); self.btnF.Location = Point(790, 126); self.btnF.Click += self.pick_rfAs
        self._family_paths = []

        # Buttons
        self.btnRun = Button(Text="Run"); self.btnRun.Location = Point(300, 190); self.btnRun.Click += self.run
        self.btnClose = Button(Text="Close"); self.btnClose.Location = Point(370, 190); self.btnClose.Click += self.close

        self.Controls.AddRange(DotNetArray[Control]([
            self.lblX,self.txtX,self.btnX,
            self.lblS,self.txtS,self.btnS,
            self.lblC,self.cboC,
            self.lblF,self.txtF,self.btnF,
            self.btnRun,self.btnClose
        ]))

        # Prefill
        sd = os.path.dirname(__file__) if '__file__' in globals() else os.getcwd()
        self.txtX.Text = find_default(sd, ".xlsx", pref=("config.xlsx","params.xlsx","parameters.xlsx")) or ""
        self.txtS.Text = find_default_shared(sd) or ""

    def pick_xlsx(self, s, a):
        dlg = OpenFileDialog(); dlg.Title="Select Excel"; dlg.Filter="Excel Workbook (*.xlsx)|*.xlsx|All Files (*.*)|*.*"; dlg.CheckFileExists=True
        if dlg.ShowDialog()==DialogResult.OK: self.txtX.Text = dlg.FileName

    def pick_sp(self, s, a):
        dlg = OpenFileDialog(); dlg.Title="Select Shared Params"; dlg.Filter="Shared Params (*.txt)|*.txt|All Files (*.*)|*.*"; dlg.CheckFileExists=True
        if dlg.ShowDialog()==DialogResult.OK: self.txtS.Text = dlg.FileName

    def pick_rfAs(self, s, a):
        dlg = OpenFileDialog(); dlg.Title="Select Families (.rfa)"; dlg.Filter="Revit Family (*.rfa)|*.rfa|All Files (*.*)|*.*"
        dlg.CheckFileExists=True; dlg.Multiselect=True
        if dlg.ShowDialog()==DialogResult.OK:
            self._family_paths = list(dlg.FileNames)
            if self._family_paths:
                shown = self._family_paths[0]
                if len(self._family_paths) > 1:
                    shown += "  (+{} more)".format(len(self._family_paths) - 1)
                self.txtF.Text = shown
            else:
                self.txtF.Text = ""

    def run(self, s, a):
        global POWER_CONN_NOT_PARAM
        POWER_CONN_NOT_PARAM = []

        xlsx, sp = self.txtX.Text.strip(), self.txtS.Text.strip()
        n = int(self.cboC.SelectedItem) if self.cboC.SelectedItem else 1

        if not (os.path.isfile(xlsx) and os.path.isfile(sp)):
            log("[ERROR] Pick valid .xlsx and .txt."); return

        try:
            rows = read_xlsx(xlsx)
            rows = filter_by_circuits(rows, n)  # <-- uses new include-EL>5 rule
            sp_byname, sp_byguid, sp_orig = load_sharedparams(app, sp)
        except Exception as e:
            log("[ERROR] Prep failed: {}".format(e)); return

        dlg_token = _attach_dialog_silencer()

        try:
            if self._family_paths:
                results = process_multiple_families(self._family_paths, rows, sp_byname, sp_byguid, sp_orig)
                pretty_report_multi(results)
            else:
                result = process_active_family(rows, sp_byname, sp_byguid, sp_orig)
                pretty_report(result)
        finally:
            _detach_dialog_silencer(dlg_token)

        _dump_power_conn_missing()

    def close(self, s, a): self.Close()

# ------------------------------ Formula helpers ------------------------------
_prefix_re_apparent = re.compile(r'^\d{2}-\d{3}(?:-\d+)?\s+')

def _basename(nm):
    return _prefix_re_apparent.sub('', (nm or '')).strip()

def _looks_like_apparent_load(nm):
    s = _basename(nm).lower()
    return ("apparent" in s) or ("kva" in s) or ("apparent load" in s)

def rewrite_apparent_if_needed(pname, formula):
    if not formula or "sqrt" not in formula.lower(): return formula
    m=re.search(r'(?i)^circuit\s*(\d+)\s+apparent\s+load(?:_ced)?$', (pname or '').strip())
    if not m: return formula
    n=m.group(1)
    phase="Circuit {} Phase_CED".format(n)
    safe="( if({} = 3, 1.73205080757, 1) )".format(phase)
    return re.sub(r'(?i)/\s*sqrt\s*\(\s*'+re.escape(phase)+r'\s*\)', r' * '+safe, formula)

def try_set_formula(fm, fp, name, formula, failed):
    try:
        target_formula = formula or ""

        if _looks_like_apparent_load(name):
            pwr_param = get_param(fm, "Power Connection Required")
            if pwr_param and (bool(pwr_param.IsInstance) == bool(fp.IsInstance)):
                target_formula = (
                    "if(Power Connection Required,"
                    "(Circuit 1 Voltage_CED * Circuit 1 FLA_CED) * (if(Circuit 1 Phase_CED = 3, 1.732051, 1)),"
                    "0)"
                )
            else:
                target_formula = "(Circuit 1 Voltage_CED * Circuit 1 FLA_CED) * (if(Circuit 1 Phase_CED = 3, 1.732051, 1))"

        if not target_formula:
            return False

        target_formula = rewrite_apparent_if_needed(name, target_formula)
        fm.SetFormula(fp, target_formula)
        return True

    except Exception as e:
        failed.append("{} ({})".format(name, e))
        return False

# ------------------------------ Excel helpers ------------------------------
def find_default(folder, ext, pref=()):
    try:
        cands = [os.path.join(folder,f) for f in os.listdir(folder) if f.lower().endswith(ext)]
        if not cands: return None
        if pref:
            for name in pref:
                for p in cands:
                    if os.path.basename(p).lower()==name: return p
        return cands[0]
    except: return None

def find_default_shared(folder):
    try:
        cands = sorted([os.path.join(folder,f) for f in os.listdir(folder) if f.lower().endswith(".txt")],
                       key=lambda p: ("shared" not in os.path.basename(p).lower(), os.path.basename(p).lower()))
        return cands[0] if cands else None
    except: return None

# ------------------------------ SharedParams ------------------------------
def open_shared(app, sp_path):
    orig = app.SharedParametersFilename
    try:
        app.SharedParametersFilename = sp_path
        sp = app.OpenSharedParameterFile()
    except:
        app.SharedParametersFilename = orig
        raise
    if sp is None:
        app.SharedParametersFilename = orig
        raise Exception("Could not open Shared Parameters: {}".format(sp_path))
    return sp, orig

def load_sharedparams(app, sp_path):
    log("[INFO] Loading Shared Parameters: {}".format(sp_path))
    sp, orig = open_shared(app, sp_path)
    byname, byguid = {}, {}
    for g in sp.Groups:
        for d in g.Definitions:
            if isinstance(d, ExternalDefinition):
                byname[(g.Name, d.Name)] = d
                try: byguid[d.GUID] = d
                except: pass
    return byname, byguid, orig

# ------------------------------ COM Excel (no external libs) ------------------------------
def _args_array(*args):
    return System.Array[System.Object](list(args))
def _set(obj, prop, val): obj.GetType().InvokeMember(prop, BindingFlags.SetProperty, None, obj, _args_array(val))
def _get(obj, prop):
    try: return obj.GetType().InvokeMember(prop, BindingFlags.GetProperty, None, obj, None)
    except: return None
def _call(obj, name, *args):
    t=obj.GetType()
    try: return t.InvokeMember(name, BindingFlags.InvokeMethod, None, obj, _args_array(*args))
    except:
        try: return t.InvokeMember(name, BindingFlags.GetProperty, None, obj, _args_array(*args) if args else None)
        except: return None
def _cell(cells, r, c):
    it = _call(cells,"Item",r,c); v=_get(it,"Value2"); return ("" if v is None else str(v)).strip()

def read_xlsx(path):
    log("[INFO] Reading Excel: {}".format(path))
    xl=wb=ws=used=cells=rows_prop=cols_prop=None
    rows=[]
    try:
        t=Type.GetTypeFromProgID("Excel.Application")
        if t is None: raise Exception("Excel not registered")
        xl = Activator.CreateInstance(t); _set(xl,"Visible",False); _set(xl,"DisplayAlerts",False)
        wb = _call(_get(xl,"Workbooks"),"Open",path)
        ws = _call(_get(wb,"Worksheets"),"Item",1)
        used = _get(ws,"UsedRange"); cells=_get(used,"Cells")
        rows_prop=_get(used,"Rows"); cols_prop=_get(used,"Columns")
        nrows=int(_get(rows_prop,"Count") or 0); ncols=int(_get(cols_prop,"Count") or 0)
        headers=[_cell(cells,1,c) for c in range(1,ncols+1)]
        missing=set(REQ_HEADERS)-set([h for h in headers if h])
        if missing: raise Exception("Missing headers: {}".format(", ".join(sorted(missing))))
        col = {h: headers.index(h)+1 for h in REQ_HEADERS}
        col_guid = headers.index(HDR_GUID)+1 if HDR_GUID in headers else None
        col_famsh = headers.index(HDR_FAMSH)+1 if HDR_FAMSH in headers else None
        col_circuit = headers.index(HDR_CIRCUIT)+1 if HDR_CIRCUIT in headers else None
        col_sort = headers.index(HDR_SORT)+1 if HDR_SORT in headers else None

        for r in range(2, nrows+1):
            name=_cell(cells,r,col["Parameter Name"])
            if not name: continue
            row = {
                "Name": name,
                "SpecType": _cell(cells,r,col["Type of Parameter"]),
                "GroupUnder": _cell(cells,r,col["Group Under"]),
                "InstanceType": _cell(cells,r,col["Instance/Type"]),
                "Formula": _cell(cells,r,col["Formula"]),
                "Guid": _cell(cells,r,col_guid) if col_guid else "",
                "FamOrShared": _cell(cells,r,col_famsh) if col_famsh else "",
                "CircuitGroup": _cell(cells,r,col_circuit) if col_circuit else "",
                "SortOrder": _cell(cells,r,col_sort) if col_sort else "",
            }
            row["CircuitGroupNum"] = parse_int(row["CircuitGroup"])
            row["SortOrderNum"] = parse_int(row["SortOrder"])
            rows.append(row)
    finally:
        try:
            if wb: _call(wb,"Close",False)
            if xl: _call(xl,"Quit")
        except: pass
        try:
            if ws: Marshal.ReleaseComObject(ws)
            if wb: Marshal.ReleaseComObject(wb)
            if xl: Marshal.ReleaseComObject(xl)
        except: pass
    return rows

def parse_int(s):
    try: return int(float((s or "").strip()))
    except: return None

# ------------------------------ CHANGED: include EL > N ------------------------------
def filter_by_circuits(rows, n):
    """Include:
       - Any row whose Circuit Group is blank or <= n (legacy behavior), AND
       - All rows whose Group Under resolves to Electrical Loads (regardless of Circuit Group),
         so Electric Heat (and other EL rows) always get processed.
    """
    out = []
    el_gid = group_type_id("Electrical Loads")
    for r in rows:
        cg = r.get("CircuitGroupNum")
        grp = group_type_id(r.get("Group Under"))

        # Always include Electrical Loads rows
        if grp == el_gid:
            out.append(r)
            continue

        # Legacy behavior for everything else
        if cg is None or (isinstance(cg, int) and 1 <= cg <= n):
            out.append(r)

    log("[INFO] Rows after circuit filter (<= {} plus ALL Electrical Loads): {}".format(n, len(out)))
    return out# ------------------------------ /CHANGED ------------------------------

# ------------------------------ Warnings Swallower ------------------------------
class _SwallowWarnings(IFailuresPreprocessor):
    def PreprocessFailures(self, accessor):
        for msg in list(accessor.GetFailureMessages()):
            try:
                if msg.GetSeverity() == FailureSeverity.Warning:
                    accessor.DeleteWarning(msg)
            except:
                pass
        return FailureProcessingResult.Continue

def _with_silent_failures(doc, txname):
    t = Transaction(doc, txname); t.Start()
    try:
        opts = t.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(_SwallowWarnings())
        opts.SetClearAfterRollback(True)
        t.SetFailureHandlingOptions(opts)
    except:
        pass
    return t

# ------------------------------ Param helpers & add/update ------------------------------
def _strip_prefix_local(nm):
    return _prefix_re_local.sub('', (nm or '')).strip()

def apparent_targets(r):
    nm = (r.get("Parameter Name") or r.get("Name") or "").strip()
    if not nm:
        return []
    base = _strip_prefix_local(nm)
    return [nm] if base == nm else [nm, base]

def parse_guid(s):
    try: return Guid(s) if s else None
    except: return None

def is_instance(s):
    return (s or "").strip().lower() in ("instance","i","inst")

def add_and_update_params(fam_doc, rows, sp_byname, sp_byguid):
    fm=fam_doc.FamilyManager
    added, skipped, f_set, f_failed, add_failed = [], [], [], [], []
    t=_with_silent_failures(fam_doc, "Add/Update Params")
    try:
        for r in rows:
            spec=resolve_spec(r.get("SpecType"))
            group=group_type_id(r.get("GroupUnder"))
            inst=is_instance(r.get("InstanceType"))
            formula=r.get("Formula") or ""
            name_target = r.get("Name") or ""
            if not name_target: continue

            fp=get_param(fm, name_target)
            if fp:
                skipped.append(name_target)
                if try_set_formula(fm, fp, name_target, formula, f_failed): f_set.append(name_target)
                continue

            famsh = (r.get("FamOrShared") or "").strip().lower()
	    is_family = (famsh == "" or famsh.startswith("family"))  # BLANK ⇒ FAMILY (new default)
	    if is_family:
                try:
                    fp=fm.AddParameter(name_target, group, spec, inst); added.append(name_target)
                    if try_set_formula(fm, fp, name_target, formula, f_failed): f_set.append(name_target)
                except Exception as e:
                    add_failed.append("{} (family add failed: {})".format(name_target, e))
                continue

            ext=None; g=parse_guid(r.get("Guid") or "")
            if g and g in sp_byguid: ext=sp_byguid[g]
            else:
                for (grp,nm),d in sp_byname.items():
                    if nm==name_target: ext=d; break
            if not ext:
                add_failed.append("{} (shared add skipped: no def)".format(name_target)); continue

            try:
                fp=fm.AddParameter(ext, group, inst); added.append(name_target)
                if try_set_formula(fm, fp, name_target, formula, f_failed): f_set.append(name_target)
            except Exception as e:
                add_failed.append("{} (shared add failed: {})".format(name_target, e))
        t.Commit()
    except:
        try: t.RollBack()
        except: pass
        raise
    return added, skipped, f_set, f_failed, add_failed

# ------------------------------ Reordering (SEPARATE TX) ------------------------------
def _get_group_type_id(fp):
    try:
        return fp.Definition.GetGroupTypeId()
    except:
        try:
            return None
        except:
            return None

def reorder_params(fam_doc, rows):
    fm = fam_doc.FamilyManager

    prefs = {}
    for r in rows:
        cg = r.get("CircuitGroupNum")
        so = r.get("SortOrderNum")
        grp = group_type_id(r.get("GroupUnder"))
        for nm in apparent_targets(r):
            prefs[nm] = {'cg': cg, 'so': so, 'grp': grp}

    fam_params = list(fm.GetParameters())
    base_index = {p.Definition.Name: i for i, p in enumerate(fam_params)}

    ELECT_LOADS_GRP = group_type_id("Electrical Loads")
    BIG = 10 ** 9

    def resolve_pref(name):
        p = prefs.get(name)
        if p is None:
            p = prefs.get(_strip_prefix_local(name))
        return p

    elect_num = []
    elect_alpha = []
    others = []

    for fp in fam_params:
        name = fp.Definition.Name
        gid = _get_group_type_id(fp)
        pref = resolve_pref(name)
        is_elec_loads = (gid == ELECT_LOADS_GRP)

        if is_elec_loads:
            cg = pref.get('cg') if pref else None
            so = pref.get('so') if pref else None
            base_name = _strip_prefix_local(name).lower()

            if isinstance(cg, int) and 1 <= cg <= 5:
                elect_num.append((fp, cg, so if isinstance(so, int) else BIG, base_name))
            else:
                elect_alpha.append((fp, base_name))
        else:
            others.append((fp, base_index.get(name, BIG)))

    elect_num_sorted   = sorted(elect_num,   key=lambda t: (t[1], t[2], t[3]))
    elect_alpha_sorted = sorted(elect_alpha, key=lambda t: t[1])
    others_sorted      = sorted(others,      key=lambda t: t[1])

    sorted_params = [t[0] for t in elect_num_sorted] + \
                    [t[0] for t in elect_alpha_sorted] + \
                    [t[0] for t in others_sorted]

    if not sorted_params or len(sorted_params) != len(fam_params):
        log("[WARN] Reorder skipped (size mismatch or empty).")
        return

    t = Transaction(fam_doc, "Reorder Parameters (Electrical Loads rule)")
    t.Start()
    try:
        fm.ReorderParameters(sorted_params)
        log("[INFO] Reorder complete: Electrical Loads (CG 1–5 by cg/so; >5 alphabetically) then others.")
        t.Commit()
    except Exception as e:
        t.RollBack()
        log("[ERROR] Reorder failed: {}".format(e))

# ------------------------------ Silent open + dialog silencer ------------------------------
def silent_open_family(path):
    mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(path)
    opts = OpenOptions(); opts.Audit = False
    return app.OpenDocumentFile(mp, opts)

def _attach_dialog_silencer():
    def _handler(sender, e):
        try:
            if e.GetType().Name == 'TaskDialogShowingEventArgs':
                msg = ""
                try: msg = e.Message or ""
                except: pass
                did = ""
                try: did = e.DialogId or ""
                except: pass
                if ("document opened" in msg.lower()) or ("pyrevit" in msg.lower()) or ("pyrevit" in did.lower()):
                    try:
                        e.OverrideResult(int(TaskDialogResult.Ok))
                    except:
                        try: e.OverrideResult(1)
                        except: pass
        except:
            pass
    try:
        __revit__.DialogBoxShowing += _handler
    except:
        pass
    return _handler

def _detach_dialog_silencer(handler):
    try:
        __revit__.DialogBoxShowing -= handler
    except:
        pass

# ------------------------------ Processing (Active + Multiple) ------------------------------
def process_active_family(rows, sp_byname, sp_byguid, sp_orig):
    fam_doc=uidoc.Document
    if fam_doc is None or not fam_doc.IsFamilyDocument:
        log("[ERROR] Active doc is not a family."); return []
    fam_path = fam_doc.PathName if fam_doc.PathName else fam_doc.Title+".rfa"
    log("[INFO] Processing ACTIVE family: {}".format(fam_path))

    if not has_param(fam_doc, "Power Connection Required"):
        POWER_CONN_NOT_PARAM.append(os.path.basename(fam_path))

    try:
        added, skipped, f_set, f_failed, add_failed = add_and_update_params(fam_doc, rows, sp_byname, sp_byguid)
        reorder_params(fam_doc, rows)
        return [(fam_path, added, skipped, f_set, f_failed, add_failed, "No save (per request)")]
    finally:
        try: app.SharedParametersFilename = sp_orig
        except: pass

def process_family_file(path, rows, sp_byname, sp_byguid):
    log("[INFO] Opening family: {}".format(path))
    fam_doc = None
    try:
        fam_doc = silent_open_family(path)
        if not fam_doc.IsFamilyDocument:
            log("[WARN] Skipped (not a family): {}".format(path))
            return (path, [], [], [], [], ["Not a family document"], "Skipped")

        if not has_param(fam_doc, "Power Connection Required"):
            POWER_CONN_NOT_PARAM.append(os.path.basename(path))

        added, skipped, f_set, f_failed, add_failed = add_and_update_params(fam_doc, rows, sp_byname, sp_byguid)
        reorder_params(fam_doc, rows)
        
        if AUTO_SAVE_NONACTIVE:
            _save_family_doc(fam_doc, path)
        
        status = "Processed{}".format("; saved" if AUTO_SAVE_NONACTIVE else "; closed without saving")
        return (path, added, skipped, f_set, f_failed, add_failed, status)
    
    except Exception as e:
        return (path, [], [], [], [], ["Error: {}".format(e)], "Failed")
    finally:
        try:
            if fam_doc: fam_doc.Close(False)
        except: pass

def process_multiple_families(paths, rows, sp_byname, sp_byguid, sp_orig):
    results=[]
    try:
        for p in paths:
            results.append(process_family_file(p, rows, sp_byname, sp_byguid))
        return results
    finally:
        try: app.SharedParametersFilename = sp_orig
        except: pass

# ------------------------------ Reporting ------------------------------
def pretty_report(report):
    if not report: log("No report."); return
    fam, added, skipped, f_set, f_failed, add_failed, status = report[0]
    log("\n----- Report (Active Family) -----")
    log("Family: {}".format(fam))
    log("  Added: {}".format(", ".join(added) if added else "(none)"))
    log("  Skipped: {}".format(", ".join(skipped) if skipped else "(none)"))
    log("  Failed adds: {}".format(", ".join(add_failed) if add_failed else "(none)"))
    log("  Formulas set: {}".format(", ".join(f_set) if f_set else "(none)"))
    log("  Formulas failed: {}".format(", ".join(f_failed) if f_failed else "(none)"))
    log("  {}".format(status))

def pretty_report_multi(results):
    if not results: log("No report."); return
    log("\n===== Multi-Family Report =====")
    for (fam, added, skipped, f_set, f_failed, add_failed, status) in results:
        log("\n-- {}".format(fam))
        log("  Status: {}".format(status))
        log("  Added: {}".format(", ".join(added) if added else "(none)"))
        log("  Skipped: {}".format(", ".join(skipped) if skipped else "(none)"))
        log("  Failed adds: {}".format(", ".join(add_failed) if add_failed else "(none)"))
        log("  Formulas set: {}".format(", ".join(f_set) if f_set else "(none)"))
        log("  Formulas failed: {}".format(", ".join(f_failed) if f_failed else "(none)"))

# ------------------------------ Entry ------------------------------
if __name__ == "__main__":
    runner = Runner()
    runner.ShowDialog()   # blocks until the user closes the window

    # After main work completes, run the helper on the same selected families
    #try:
    #    family_paths = getattr(runner, "_family_paths", []) or []
    #    if family_paths:
    #        print("[POST] Deleting Power Balance connector(s) and saving {} file(s)…".format(len(family_paths)))
    #        run_helper_delete_power_balance_and_save(family_paths)
    #    else:
    #        print("[POST] No selected families; helper skipped.")
    #except Exception as e:
    #    print("[ERROR] Post-helper failed: {}".format(e))
