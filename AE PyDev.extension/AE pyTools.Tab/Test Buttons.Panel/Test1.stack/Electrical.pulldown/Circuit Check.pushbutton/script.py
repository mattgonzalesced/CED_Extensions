# -*- coding: utf-8 -*-
"""
pyRevit | Toggle Un-/Circuited Fixture Check
• If red/green filters don’t exist → create & enable them
• If they exist and are enabled    → disable them
• If they exist and are disabled   → enable them
No Temporary-Hide/Isolate.  No view-template gymnastics.
IronPython 2.7 compliant.
FIX: Newly created filters are now explicitly disabled so the first run no longer shows the
     misleading "Fixture check OFF" message.
"""
import clr, sys, traceback
from pyrevit import revit, DB
from Autodesk.Revit.DB import (
    BuiltInCategory, BuiltInParameter, Transaction, ElementId,
    FilteredElementCollector, ParameterFilterRuleFactory,
    ElementParameterFilter, ParameterFilterElement,
    OverrideGraphicSettings, Color
)

# ── Small UI helper ───────────────────────────────────────────────
try:
    from Autodesk.Revit.UI import TaskDialog
    def alert(msg, title="Fixture Check"):
        TaskDialog.Show(title, msg)
except ImportError:
    def alert(msg, title="Fixture Check"):
        print("[{0}] {1}".format(title, msg))

doc  = revit.doc
view = doc.ActiveView

UNC_NAME = "Uncircuited Fixtures"
CIR_NAME = "Circuited Fixtures"

# ── helpers ───────────────────────────────────────────────────────

def pick_parameter(sample):
    """Return (paramId, 'int'|'str') or (None,None)"""
    for bip, kind in (
            (BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER, 'int'),
            (BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM, 'str'),
            (BuiltInParameter.RBS_SYSTEM_NAME_PARAM, 'str')):
        p = sample.get_Parameter(bip)
        if p and ((kind == 'int' and p.StorageType == DB.StorageType.Integer) or
                  (kind == 'str' and p.StorageType == DB.StorageType.String)):
            return ElementId(bip), kind
    return None, None


def ensure_filter(name, rule_list, cats):
    """Create or update a ParameterFilterElement; ensure on view; return id"""
    epf = ElementParameterFilter(rule_list)
    existing = next(
        (f for f in FilteredElementCollector(doc)
              .OfClass(ParameterFilterElement).ToElements()
         if f.Name == name), None)
    if existing:
        existing.SetElementFilter(epf)
        existing.SetCategories(cats)   # extend categories if needed
        fid = existing.Id
    else:
        fid = ParameterFilterElement.Create(doc, name, cats, epf).Id

    if fid not in view.GetFilters():
        view.AddFilter(fid)
        # NEW ➜ make sure a freshly-added filter starts DISABLED so our
        #       first toggle correctly reports "ON" instead of "OFF".
        view.SetIsFilterEnabled(fid, False)

    return fid

# ── grab category list & a sample fixture ─────────────────────────
from System.Collections.Generic import List

cats = List[ElementId]([
    doc.Settings.Categories.get_Item(BuiltInCategory.OST_ElectricalFixtures).Id,
    doc.Settings.Categories.get_Item(BuiltInCategory.OST_LightingFixtures).Id
])

sample = next((e for e in FilteredElementCollector(doc)
                     .OfCategory(BuiltInCategory.OST_ElectricalFixtures)
                     .WhereElementIsNotElementType()), None)
if not sample:
    alert("No Electrical Fixtures found – nothing to do.")
    sys.exit()

param_id, ptype = pick_parameter(sample)
if not param_id:
    alert("Fixtures lack Circuit-Number / Panel / System-Name parameters.")
    sys.exit()

# build rules (one empty, one non-empty)
if ptype == 'int':
    rule_empty    = ParameterFilterRuleFactory.CreateEqualsRule(param_id, 0)
    rule_filled   = ParameterFilterRuleFactory.CreateNotEqualsRule(param_id, 0)
else:
    rule_empty    = ParameterFilterRuleFactory.CreateEqualsRule(param_id, "")
    rule_filled   = ParameterFilterRuleFactory.CreateNotEqualsRule(param_id, "")
rl_empty  = List[DB.FilterRule](); rl_empty.Add(rule_empty)
rl_filled = List[DB.FilterRule](); rl_filled.Add(rule_filled)

# colours
RED, GREEN = Color(255,60,60), Color(60,200,60)

# ── TRANSACTION ──────────────────────────────────────────────────

t = Transaction(doc, "Toggle Fixture Check")
try:
    t.Start()

    # 1) make sure filters exist & attached (ensure_filter now disables new ones)
    fid_unc = ensure_filter(UNC_NAME, rl_empty,  cats)
    fid_cir = ensure_filter(CIR_NAME, rl_filled, cats)

    # 2) are they currently enabled?
    already_on = view.GetIsFilterEnabled(fid_unc)

    if already_on:
        # ---- TURN OFF ----
        view.SetIsFilterEnabled(fid_unc, False)
        view.SetIsFilterEnabled(fid_cir, False)
        msg = "Fixture check OFF – red/green filters disabled."
    else:
        # ---- TURN ON (also set colours) ----
        ogs_r = OverrideGraphicSettings(); ogs_r.SetProjectionLineColor(RED)
        ogs_g = OverrideGraphicSettings(); ogs_g.SetProjectionLineColor(GREEN)

        view.SetFilterOverrides(fid_unc, ogs_r)
        view.SetFilterOverrides(fid_cir, ogs_g)

        view.SetIsFilterEnabled(fid_unc, True)
        view.SetIsFilterEnabled(fid_cir, True)
        msg = "Fixture check ON.\n\n• Un-circuited = RED\n• Circuited = GREEN"

    t.Commit()
    alert(msg)

except Exception as ex:
    if t.HasStarted():
        t.RollBack()
    alert("Error (rolled back).\n\n{0}\n{1}".format(
        ex, traceback.format_exc()), title="Fixture Check – Error")