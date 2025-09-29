# -*- coding: utf-8 -*-
# Revit 2025+ | pyRevit IronPython | Active Family Only | No Save
# Features:
# - UI: pick Excel + Shared Params, circuits 1–5
# - Excel: reads required columns; filters rows by "Circuit Group"
# - Adds Family/Shared parameters; sets formulas (incl. Apparent Load safe-rewrite)
# - Reorders by Circuit Group then Sort Order (not alphabetical)
# - Does NOT create any new families, and does NOT save

import clr
import os
import re
from System import Type, Activator, Array, Object, Guid
from System.Reflection import BindingFlags
from System.Runtime.InteropServices import Marshal

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
clr.AddReference('RevitServices')

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import (Form, Label, TextBox, Button, ComboBox, ComboBoxStyle,
                                  OpenFileDialog, Control, DialogResult)
from System.Drawing import Point, Size
from System import Array as DotNetArray

uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

# ------------------------------ Config / constants ------------------------------
REQ_HEADERS = ["Parameter Name", "Type of Parameter", "Group Under", "Instance/Type", "Formula"]
HDR_GUID, HDR_FAMSH = "GUID", "Family/Shared"
HDR_CIRCUIT, HDR_SORT = "Circuit Group", "Sort Order"

GROUP_MAP = {
    "Text":"Text","Constraints":"Constraints","Geometry":"Geometry","Dimensions":"Geometry",
    "Identity Data":"IdentityData","Materials and Finishes":"Materials","Construction":"Construction",
    "Data":"Data","Graphics":"Graphics","Other":"Other",
    "PG_TEXT":"Text","PG_CONSTRAINTS":"Constraints","PG_GEOMETRY":"Geometry",
    "PG_IDENTITY_DATA":"IdentityData","PG_MATERIALS":"Materials","PG_CONSTRUCTION":"Construction",
    "PG_DATA":"Data","PG_GRAPHICS":"Graphics",
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

# ------------------------------ UI ------------------------------
class Runner(Form):
    def __init__(self):
        self.Text = "Add Parameters (Active Family Only)"
        self.Size = Size(880, 210)

        self.lblX = Label(Text="Excel"); self.lblX.Location = Point(10, 20)
        self.txtX = TextBox(); self.txtX.Location = Point(140, 18); self.txtX.Size = Size(640, 22)
        self.btnX = Button(Text="Browse"); self.btnX.Location = Point(790, 16); self.btnX.Click += self.pick_xlsx

        self.lblS = Label(Text="Shared Params"); self.lblS.Location = Point(10, 55)
        self.txtS = TextBox(); self.txtS.Location = Point(140, 53); self.txtS.Size = Size(640, 22)
        self.btnS = Button(Text="Browse"); self.btnS.Location = Point(790, 51); self.btnS.Click += self.pick_sp

        self.lblC = Label(Text="Circuits (1–5):"); self.lblC.Location = Point(10, 90)
        self.cboC = ComboBox(); self.cboC.Location = Point(140, 88); self.cboC.DropDownStyle = ComboBoxStyle.DropDownList
        for i in range(1,6): self.cboC.Items.Add(str(i))
        self.cboC.SelectedIndex = 0

        self.btnRun = Button(Text="Run"); self.btnRun.Location = Point(300, 86); self.btnRun.Click += self.run
        self.btnClose = Button(Text="Close"); self.btnClose.Location = Point(370, 86); self.btnClose.Click += self.close

        self.Controls.AddRange(DotNetArray[Control]([self.lblX,self.txtX,self.btnX,self.lblS,self.txtS,self.btnS,self.lblC,self.cboC,self.btnRun,self.btnClose]))
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

    def run(self, s, a):
        xlsx, sp = self.txtX.Text.strip(), self.txtS.Text.strip()
        n = int(self.cboC.SelectedItem) if self.cboC.SelectedItem else 1
        fam_doc = uidoc.Document
        if fam_doc is None or not fam_doc.IsFamilyDocument: log("[ERROR] Active doc is not a Family."); return
        if not (os.path.isfile(xlsx) and os.path.isfile(sp)): log("[ERROR] Pick valid .xlsx and .txt."); return

        try:
            rows = read_xlsx(xlsx)
            rows = filter_by_circuits(rows, n)
            sp_byname, sp_byguid, sp_orig = load_sharedparams(app, sp)
            report = process_active_family(rows, sp_byname, sp_byguid, sp_orig)
            pretty_report(report)
        except Exception as e:
            log("[ERROR] {}".format(e))
        finally:
            pass

    def close(self, s, a): self.Close()

# ------------------------------ Helpers: files/Excel/SharedParams ------------------------------
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
        for p in cands:
            try: _ = open_shared(app, p); return p
            except: continue
        return None
    except: return None

def open_shared(app, sp_path):
    orig = app.SharedParametersFilename; app.SharedParametersFilename = sp_path
    sp = app.OpenSharedParameterFile()
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
    log("[INFO] Shared defs: {} ({} with GUID)".format(len(byname), len(byguid)))
    return byname, byguid, orig

# Excel via COM (first sheet)
def _args_array(*items): return Array[Object](list(items))
def _set(obj, prop, val):
    try: obj.GetType().InvokeMember(prop, BindingFlags.SetProperty, None, obj, _args_array(val))
    except: pass
def _get(obj, prop):
    try: return obj.GetType().InvokeMember(prop, BindingFlags.GetProperty, None, obj, None)
    except: return None
def _call(obj, name, *args):
    t=obj.GetType()
    try: return t.InvokeMember(name, BindingFlags.GetProperty, None, obj, _args_array(*args))
    except:
        try: return t.InvokeMember(name, BindingFlags.InvokeMethod, None, obj, _args_array(*args))
        except: return None
def _cell(cells, r, c):
    it = _call(cells,"Item",r,c); v=_get(it,"Value2"); return ("" if v is None else str(v)).strip()

def read_xlsx(path):
    log("[INFO] Reading Excel: {}".format(path))
    xl=wb=ws=used=cells=rows_prop=cols_prop=None
    rows=[]
    try:
        t=Type.GetTypeFromProgID("Excel.Application");
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
        col_fsh  = headers.index(HDR_FAMSH)+1 if HDR_FAMSH in headers else None
        col_cg   = headers.index(HDR_CIRCUIT)+1 if HDR_CIRCUIT in headers else None
        col_so   = headers.index(HDR_SORT)+1    if HDR_SORT in headers else None

        def is_instance(val):
            s=(str(val or "")).strip().lower()
            if s=="type": return False
            return s in ("instance","true","yes","1","y","t")

        def norm_formula(f):
            if not f: return ""
            s=str(f).strip()
            if s.startswith("="): s=s[1:].lstrip()
            return s.replace(u"\u2212","-").strip()

        for r in range(2, nrows+1):
            pname=_cell(cells,r,col["Parameter Name"])
            if not pname: continue
            rows.append({
                "Name": pname,
                "TypeInfo": _cell(cells,r,col["Type of Parameter"]),
                "GroupUnder": _cell(cells,r,col["Group Under"]),
                "IsInstance": is_instance(_cell(cells,r,col["Instance/Type"])),
                "Formula": norm_formula(_cell(cells,r,col["Formula"])),
                "Guid": _cell(cells,r,col_guid) if col_guid else "",
                "FamilyOrShared": (_cell(cells,r,col_fsh) if col_fsh else "Shared").strip(),
                "CircuitGroupRaw": _cell(cells,r,col_cg) if col_cg else "",
                "CircuitGroupNum": parse_int(_cell(cells,r,col_cg)) if col_cg else None,
                "SortOrderNum": parse_int(_cell(cells,r,col_so)) if col_so else None
            })
        log("[INFO] Parsed {} rows.".format(len(rows)))
        return rows
    finally:
        try:
            if wb is not None: _call(wb,"Close",False)
            if xl is not None: _call(xl,"Quit")
        except: pass
        for o in (cols_prop, rows_prop, cells, used, ws, wb, xl):
            try:
                if o is not None: Marshal.ReleaseComObject(o)
            except: pass

def parse_int(val):
    if val is None: return None
    m=re.search(r'(-?\d+)', str(val).strip())
    return int(m.group(1)) if m else None

def filter_by_circuits(rows, n):
    allowed=set(range(1,n+1))
    anynum=any(r.get("CircuitGroupNum") is not None for r in rows)
    if not anynum:
        log("[INFO] No numeric '{}' values; nothing to do.".format(HDR_CIRCUIT))
        return []
    kept=[r for r in rows if r.get("CircuitGroupNum") in allowed]
    log("[INFO] Circuit filter: kept {} (of {}).".format(len(kept), len(rows)))
    return kept

# ------------------------------ Revit helpers ------------------------------
_GROUP_LABEL_MAP = None

def _norm_key(s):
    # normalize user/Excel text like "Electrical - Loads" or "PG_ELECTRICAL_LOADS"
    if not s: return ""
    s = s.strip()
    # Allow both exact label and a simplified key for fuzzy matches
    return "".join(ch for ch in s.lower() if ch.isalnum())

def _build_group_label_map():
    """
    Build a dict mapping:
      - exact Revit label (lowercased) -> GroupTypeId
      - simplified key (letters/digits only) -> GroupTypeId
      - property name variants (e.g., 'PGELECTRICALLOADS') -> GroupTypeId
    Uses LabelUtils so it follows whatever Revit shows in the UI ("Electrical - Loads", etc).
    """
    global _GROUP_LABEL_MAP
    if _GROUP_LABEL_MAP is not None:
        return _GROUP_LABEL_MAP

    _GROUP_LABEL_MAP = {}
    try:
        gt_type = clr.GetClrType(GroupTypeId)
        props = gt_type.GetProperties(BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy)
        for p in props:
            try:
                val = p.GetValue(None, None)
                if not isinstance(val, ForgeTypeId):
                    continue
                # Try to get the UI label for this group id.
                label = None
                try:
                    # Revit 2024+ API
                    label = LabelUtils.GetLabelForGroupTypeId(val)
                except:
                    # Fallback (older APIs) – if missing, we’ll just use property name
                    label = None

                if label:
                    _GROUP_LABEL_MAP[label.lower()] = val
                    _GROUP_LABEL_MAP[_norm_key(label)] = val
                # Also index by the raw property name (and a de-PG_ variant) for robustness
                pname = p.Name  # e.g., "ElectricalLoads", "Data", "Text"
                _GROUP_LABEL_MAP[pname.lower()] = val
                _GROUP_LABEL_MAP[_norm_key(pname)] = val
                if pname.startswith("PG_"):
                    _GROUP_LABEL_MAP[pname[3:].lower()] = val
                    _GROUP_LABEL_MAP[_norm_key(pname[3:])] = val
            except:
                continue
    except:
        pass
    return _GROUP_LABEL_MAP

def group_type_id(name):
    """
    Resolve a 'Group Under' cell value (e.g., 'Electrical - Loads') to a GroupTypeId.
    Priority:
      1) Exact UI label via LabelUtils (case-insensitive)
      2) Simplified key match (letters/digits only)
      3) Raw property-name/PG_ name match
      4) Your legacy GROUP_MAP fallback
      5) Default to Data
    """
    if not name:
        return GroupTypeId.Data

    # 1–3) Dynamic discovery from Revit
    mp = _build_group_label_map()
    key_exact = name.strip().lower()
    key_slim  = _norm_key(name)

    gt = mp.get(key_exact) or mp.get(key_slim)
    if gt:
        return gt

    # 4) Fallback to your legacy alias table (kept for safety)
    key = GROUP_MAP.get(name) or next((GROUP_MAP[k] for k in GROUP_MAP if k.lower()==key_exact), None)
    if key:
        try:
            return getattr(GroupTypeId, key, GroupTypeId.Data)
        except:
            pass

    # 5) Last resort
    return GroupTypeId.Data

def resolve_spec(type_str):
    if not type_str: return None
    s="".join(ch for ch in type_str.lower() if ch.isalnum())
    for k,v in TYPE_ALIAS.items():
        if "".join(ch for ch in k.lower() if ch.isalnum())==s:
            try: ft=getattr(ParameterTypeId, v);
            except: ft=None
            if isinstance(ft, ForgeTypeId): return ft
    for lbl in ("Url","YesNo","Text","Number","Integer","Material","Length","Area","Volume","Angle","Slope","Currency"):
        try:
            ft=getattr(ParameterTypeId, lbl)
            if isinstance(ft, ForgeTypeId): return ft
        except: pass
    return None

def get_param(fm, name):
    for p in fm.GetParameters():
        if p.Definition and p.Definition.Name==name: return p
    return None

def parse_guid(g):
    try: return Guid(g) if g else None
    except: return None

def is_instance(fp):
    try: return bool(fp.IsInstance)
    except: return False

def is_reporting(fp):
    try: return bool(fp.IsReporting)
    except: return False

def dtype_str(defn):
    try:
        dt=defn.GetDataType();
        return dt.TypeId if hasattr(dt,"TypeId") else str(dt)
    except: return "n/a"

# Apparent Load helpers
def apparent_generic(name):
    return re.sub(r'[\s_]+','', (name or '').lower()) in ("apparentload","apparentloadced")
def apparent_family_name(name):
    return bool(re.match(r'(?i)^circuit\s*\d+\s+apparent\s+load(?:_ced)?$', (name or '').strip()))
def apparent_targets(row):
    n=row.get("CircuitGroupNum")
    if apparent_generic(row.get("Name")) and isinstance(n,int):
        return ["Circuit {} Apparent Load_CED".format(n)]
    return [row.get("Name","")]

def rewrite_apparent_if_needed(pname, formula):
    if not formula or "sqrt" not in formula.lower(): return formula
    m=re.search(r'(?i)^circuit\s*(\d+)\s+apparent\s+load(?:_ced)?$', (pname or '').strip())
    if not m: return formula
    n=m.group(1)
    phase="Circuit {} Phase_CED".format(n)
    safe="( if({} = 3, 1.73205080757, 1) )".format(phase)
    return re.sub(r'(?i)/\s*sqrt\s*\(\s*'+re.escape(phase)+r'\s*\)', r' * '+safe, formula)

def try_set_formula(fm, fp, pname, formula, rec_apparent_fail):
    if not formula: return False
    if is_instance(fp) or is_reporting(fp): return False
    is_app = apparent_family_name(pname) or apparent_generic(pname)
    candidates=[formula]
    if is_app:
        sf=rewrite_apparent_if_needed(pname, formula)
        if sf and sf!=formula: candidates.append(sf)
    last_err=None
    for f in candidates:
        try:
            fm.SetFormula(fp, f)
            return True
        except Exception as e:
            last_err=e
    if is_app and rec_apparent_fail is not None:
        rec_apparent_fail.append("{} ({})".format(pname, last_err))
    return False

# ------------------------------ Core processing ------------------------------
def add_and_update_params(fam_doc, rows, sp_byname, sp_byguid):
    fm=fam_doc.FamilyManager
    added, skipped, f_set, f_failed, add_failed = [], [], [], [], []

    t = Transaction(fam_doc, "Add/Update Parameters from Excel")
    t.Start()
    try:
        for r in rows:
            oname = r["Name"]; group=group_type_id(r["GroupUnder"]); inst=r["IsInstance"]
            formula=r["Formula"]; famsh=(r.get("FamilyOrShared") or "Shared").lower()
            type_label=r.get("TypeInfo") or ""; targets=apparent_targets(r)
            special = (apparent_generic(oname) and len(targets)==1 and targets[0]!=oname)

            # First: operate on existing targets
            for tname in targets:
                fp=get_param(fm, tname)
                if fp:
                    skipped.append(tname+" (exists)")
                    if try_set_formula(fm, fp, tname, formula, f_failed): f_set.append(tname)
                else:
                    if special:
                        add_failed.append("{} → missing '{}' (not creating)".format(oname, tname))

            if special: # don't create Apparent Load targets
                continue
            if all(get_param(fm, t) is not None for t in targets):
                continue

            # Non-special creation path (by original name)
            existing = get_param(fm, oname)
            if existing:
                skipped.append(oname+" (exists)")
                if try_set_formula(fm, existing, oname, formula, f_failed): f_set.append(oname)
                continue

            if famsh.startswith("family"):
                spec=resolve_spec(type_label)
                if not isinstance(spec, ForgeTypeId):
                    add_failed.append("{} (family add skipped: bad type '{}')".format(oname, type_label)); continue
                try:
                    fp=fm.AddParameter(oname, group, spec, inst); added.append(oname)
                    if try_set_formula(fm, fp, oname, formula, f_failed): f_set.append(oname)
                except Exception as e:
                    add_failed.append("{} (family add failed: {})".format(oname, e))
                continue

            # Shared param creation
            ext=None; g=parse_guid(r.get("Guid") or "")
            if g and g in sp_byguid: ext=sp_byguid[g]
            else:
                for (grp,nm),d in sp_byname.items():
                    if nm==oname: ext=d; break
            if not ext:
                add_failed.append("{} (shared add skipped: no def)".format(oname)); continue
            try:
                fp=fm.AddParameter(ext, group, inst); added.append(oname)
                if try_set_formula(fm, fp, oname, formula, f_failed): f_set.append(oname)
            except Exception as e:
                add_failed.append("{} (shared add failed: {})".format(oname, e))
        t.Commit()
    except Exception as e:
        t.RollBack(); raise
    return added, skipped, f_set, f_failed, add_failed

def reorder_params(fam_doc, rows):
    fm=fam_doc.FamilyManager
    prefs={}
    for r in rows:
        cg=r.get("CircuitGroupNum"); so=r.get("SortOrderNum")
        for name in apparent_targets(r):
            if cg is not None or so is not None: prefs[name]=(cg,so)
    if not prefs:
        log("[INFO] No '{}' or '{}' found; skip reorder.".format(HDR_CIRCUIT, HDR_SORT));
        return
    fam_params=list(fm.GetParameters())
    base={p.Definition.Name: i for i,p in enumerate(fam_params)}
    BIG=10**9
    def key(p):
        n=p.Definition.Name; cg,so=prefs.get(n,(None,None))
        return (cg if isinstance(cg,int) else BIG, so if isinstance(so,int) else BIG, base.get(n,BIG))
    sorted_params=sorted(fam_params, key=key)

    t=Transaction(fam_doc, "Reorder Parameters (Circuit Group, Sort Order)")
    t.Start()
    try:
        fm.ReorderParameters(sorted_params)
        t.Commit(); log("[INFO] Reorder complete (by '{}' then '{}').".format(HDR_CIRCUIT, HDR_SORT))
    except Exception as e:
        t.RollBack(); log("[ERROR] Reorder failed: {}".format(e))

def process_active_family(rows, sp_byname, sp_byguid, sp_orig):
    fam_doc=uidoc.Document
    if fam_doc is None or not fam_doc.IsFamilyDocument:
        log("[ERROR] Active doc is not a family."); return []
    fam_path = fam_doc.PathName if fam_doc.PathName else fam_doc.Title+".rfa"
    log("[INFO] Processing ACTIVE family: {}".format(fam_path))

    try:
        added, skipped, f_set, f_failed, add_failed = add_and_update_params(fam_doc, rows, sp_byname, sp_byguid)
        reorder_params(fam_doc, rows)
        return [(fam_path, added, skipped, f_set, f_failed, add_failed, "No save (per request)")]
    finally:
        try:
            app.SharedParametersFilename = sp_orig
        except: pass

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

# ------------------------------ Entry ------------------------------
if __name__ == "__main__":
    Runner().ShowDialog()
