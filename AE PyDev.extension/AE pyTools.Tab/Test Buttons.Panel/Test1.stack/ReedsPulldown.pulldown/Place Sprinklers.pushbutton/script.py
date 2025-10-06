# -*- coding: utf-8 -*-
# script.py — thin orchestrator

# SETTINGS
DRY_RUN = False
ONLY_CURRENT_LEVEL = False
PREFER_SELECTION = True
VERBOSE = True

import os, sys, clr

# -------- Bind to Revit safely --------
uiapp = uidoc = doc = app = active_view = None
try:
    from pyrevit import revit
    uiapp = revit.uiapp; uidoc = revit.uidoc; doc = revit.doc
    app = uiapp.Application if uiapp else None
except Exception:
    pass
if app is None:
    try:
        uiapp = __revit__; app = uiapp.Application
        uidoc = uiapp.ActiveUIDocument; doc = uidoc.Document if uidoc else None
    except Exception:
        pass
if app is None:
    clr.AddReference('RevitServices')
    from RevitServices.Persistence import DocumentManager
    dm = DocumentManager.Instance
    uiapp = dm.CurrentUIApplication; uidoc = uiapp.ActiveUIDocument if uiapp else None
    doc = dm.CurrentDBDocument; app = uiapp.Application if uiapp else None
if app is None or doc is None:
    raise EnvironmentError("Open a project in Revit and run from a pyRevit button.")
active_view = doc.ActiveView

# -------- Paths & sys.path --------
SCRIPT_DIR = os.path.dirname(__file__)

def _find_dir(base_name):
    # 1) walk up to 4 ancestors and check each for base_name
    anc = SCRIPT_DIR
    candidates = []
    for _ in range(5):  # script -> parent -> grandparent -> great-grandparent -> great-great
        candidates.append(os.path.join(anc, base_name))
        anc = os.path.dirname(anc)

    # 2) for the script dir and its first two parents, also look inside sibling *.pushbutton folders
    for root in [SCRIPT_DIR,
                 os.path.dirname(SCRIPT_DIR),
                 os.path.dirname(os.path.dirname(SCRIPT_DIR))]:
        try:
            for name in os.listdir(root):
                if name.endswith(".pushbutton"):
                    candidates.append(os.path.join(root, name, base_name))
        except Exception:
            pass

    for p in candidates:
        if os.path.isdir(p):
            return p
    return None

# Resolve rules/
RULES_DIR = _find_dir("rules")
if not RULES_DIR:
    raise EnvironmentError("[FATAL] Couldn't find a 'rules' folder near: {0}".format(SCRIPT_DIR))
IDENTIFY = os.path.join(RULES_DIR, "identify_spaces.json")
SPRINKLERS = os.path.join(RULES_DIR, "sprinkler_rules.json")
ELEC_DIR = os.path.join(RULES_DIR, "electrical")

# Resolve lib/
LIB_DIR = _find_dir("lib")
if not LIB_DIR:
    raise EnvironmentError("[FATAL] Couldn't find a 'lib' folder near: {0}".format(SCRIPT_DIR))
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

print("[PATH] rules -> {0}".format(RULES_DIR))
print("[PATH] lib   -> {0}".format(LIB_DIR))

# -------- Imports (your existing rules_loader stays) --------
from log import log
from single_undo import run_as_single_undo
from rules_loader import build_rule_for_room, deep_merge
from electrical_loader import load_electrical_rules, get_room_profile
from placement import pick_fixture_symbol as pick_sprinkler_symbol
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
from place_by_space import (
    get_target_spaces, spatial_display_name, place_fixtures_in_space
)

def _view_level_id(view):
    try:
        gl = getattr(view, "GenLevel", None)
        return gl.Id if gl else view.LevelId
    except:
        return None
def do_all():
    spaces = get_target_spaces(doc, uidoc, active_view, only_current_level=False, prefer_selection=True)
    if VERBOSE:
        print("[INFO] {} spaces found. Processing…".format(len(spaces)))

    total_lights = total_recepts = 0

    for sp in spaces:
        nm, num = spatial_display_name(sp)

        # LIGHTS
        cat, light_rule = build_rule_for_room(nm, IDENTIFY, SPRINKLERS)
        cands = light_rule.get('fixture_candidates') or []
        eff_light = deep_merge(light_rule, cands[0]) if cands else dict(light_rule)
        planned, placed = place_fixtures_in_space(doc, active_view, sp, eff_light,
                                                  pick_sprinkler_symbol, dry_run=DRY_RUN, verbose=VERBOSE)
        total_lights += len(placed)

run_as_single_undo(doc, "Auto-Place Lights + Receptacles", do_all)

