# -*- coding: utf-8 -*-
# Deterministic Agent Orchestrator (no run() requirement on tools)

from __future__ import print_function
import os
import sys
import json
import traceback
import datetime
import runpy

from Autodesk.Revit.UI import TaskDialog

# Try pyRevit pretty output
try:
    from pyrevit import script as _py_out
    OUTPUT = _py_out.get_output()
except:
    OUTPUT = None

# ------------------------------------------------------------------------
# Path Setup: tools/lib/rules are siblings of this pushbutton folder
# ------------------------------------------------------------------------
THIS_DIR = os.path.dirname(__file__)
PANEL_DIR = os.path.dirname(THIS_DIR)  # e.g. ...\RP_Panel.panel
TOOLS_DIR = os.path.join(PANEL_DIR, "tools")
LIB_DIR   = os.path.join(PANEL_DIR, "lib")
RULES_DIR = os.path.join(PANEL_DIR, "rules")

for p in (TOOLS_DIR, LIB_DIR, RULES_DIR):
    if os.path.isdir(p) and p not in sys.path:
        sys.path.insert(0, p)

# ------------------------------------------------------------------------
# Logging helpers
# ------------------------------------------------------------------------
def log(msg):
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    line = "[Agent {0}] {1}".format(ts, msg)
    print(line)
    if OUTPUT:
        try:
            OUTPUT.print_md(line)
        except:
            pass

def info_dialog(title, body):
    try:
        TaskDialog.Show(title, body)
    except:
        log("{0}: {1}".format(title, body))

# ------------------------------------------------------------------------
# Load agent config
# ------------------------------------------------------------------------

#-------------- DEFAULT CONFIG is the order it will execute the tools --------------
DEFAULT_CONFIG = {
    "agent_name": "DeterministicMEPAgent",
    "version": "0.1.0",
    "dry_run": False,
    "stop_on_fail": True,
    "plan": [
        {"tool": "place_receptacles"},
        {"tool": "route_circuits"}
    ]
}

CONFIG_FILE = os.path.join(RULES_DIR, "agent.json")

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        except Exception as e:
            log("WARN: Failed to load agent.json: {0}".format(e))
    return DEFAULT_CONFIG

# ------------------------------------------------------------------------
# Build Revit context (if needed)
# ------------------------------------------------------------------------
def build_ctx():
    from pyrevit import revit
    return {
        "app": revit.doc.Application,
        "uidoc": revit.uidoc,
        "doc": revit.doc,
    }

# ------------------------------------------------------------------------
# Execute a tool script directly
# ------------------------------------------------------------------------
def execute_tool_script(tool_name, ctx):
    """Run tools/<tool_name>.py as if launched by a pyRevit pushbutton."""
    import os, sys, runpy
    # IronPython-compatible builtins
    try:
        import builtins as _builtins
    except ImportError:
        import __builtin__ as _builtins

    import pyrevit as _py  # import the package first

    # --- Resolve UIApplication first (without touching pyrevit.revit yet) ---
    uiapp = getattr(_py, "__revit__", None)
    try:
        # Prefer the real one if pyRevit has it
        if uiapp is None and hasattr(_py, "HOST_APP") and getattr(_py.HOST_APP, "uiapp", None):
            uiapp = _py.HOST_APP.uiapp
    except Exception:
        pass
    # Last-resort: try builtins (some environments inject it there)
    if uiapp is None and hasattr(_builtins, "__revit__"):
        uiapp = getattr(_builtins, "__revit__", None)

    if uiapp is None:
        raise EnvironmentError("Can’t resolve UIApplication (__revit__). Run from pyRevit with a project open.")

    # Make pyrevit’s package-level global available BEFORE importing pyrevit.revit
    setattr(_py, "__revit__", uiapp)

    # Now it’s safe to use pyrevit.revit helpers
    from pyrevit import revit as _revit
    from pyrevit import script as _script
    from pyrevit import forms as _forms
    try:
        from pyrevit import HOST_APP as _HOST_APP
    except Exception:
        _HOST_APP = None

    # Derive uidoc/doc
    uidoc = getattr(_revit, "uidoc", None)
    if uidoc is None and hasattr(uiapp, "ActiveUIDocument"):
        uidoc = uiapp.ActiveUIDocument
    doc = getattr(_revit, "doc", None)
    if doc is None and uidoc is not None:
        doc = getattr(uidoc, "Document", None)

    # Require a project document
    if doc is None or (hasattr(doc, "IsFamilyDocument") and doc.IsFamilyDocument):
        raise EnvironmentError("Open a project document (not a family) and run from a pyRevit button.")

    # ---- Build a pushbutton-like globals dict ----
    tool_path = os.path.join(TOOLS_DIR, tool_name + ".py")
    if not os.path.exists(tool_path):
        raise RuntimeError("Tool not found: {0}".format(tool_path))

    globals_dict = {
        "__name__": "__main__",
        "__file__": tool_path,
        "__revit__": uiapp,          # IMPORTANT: UIApplication
        "__uidoc__": uidoc,
        "__doc__": doc,
        "__context__": "project",
        "__window__": None,
        "revit": _revit,
        "script": _script,
        "forms": _forms,
        "HOST_APP": _HOST_APP,
        "CTX": ctx,
    }

    # Mirror for legacy patterns
    setattr(_builtins, "__revit__", uiapp)
    setattr(_builtins, "__uidoc__", uidoc)
    setattr(_builtins, "__doc__", doc)
    setattr(_builtins, "__context__", "project")
    setattr(_builtins, "revit", _revit)
    setattr(_builtins, "script", _script)
    setattr(_builtins, "forms", _forms)
    setattr(_builtins, "HOST_APP", _HOST_APP)

    # Env hints
    os.environ.setdefault("PYREVIT_RUNNING", "1")
    os.environ.setdefault("PYREVIT_EXEC_CTX", "project")

    # Emulate pushbutton cwd + sys.path[0]
    tool_dir = os.path.dirname(tool_path)
    old_cwd = os.getcwd()
    inserted_path0 = False
    try:
        if not sys.path or sys.path[0] != tool_dir:
            sys.path.insert(0, tool_dir)
            inserted_path0 = True
        os.chdir(tool_dir)

        log("Running tool: {0}".format(tool_name))
        log("CWD={0} | FILE={1}".format(os.getcwd(), tool_path))
        runpy.run_path(tool_path, globals_dict)
        log("Completed: {0}".format(tool_name))
    finally:
        os.chdir(old_cwd)
        if inserted_path0 and sys.path and sys.path[0] == tool_dir:
            sys.path.pop(0)
        # polite cleanup
        for k in ("__revit__", "__uidoc__", "__doc__", "__context__", "revit", "script", "forms", "HOST_APP"):
            if hasattr(_builtins, k):
                try: delattr(_builtins, k)
                except: pass

# --- Agent-side selection helpers (no tool edits needed) ---
def _collect_spaces(doc, mode="all"):
    from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
    elems = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_MEPSpaces
    ).WhereElementIsNotElementType().ToElements()
    if mode == "active_view_level":
        try:
            view = doc.ActiveView
            if hasattr(view, "GenLevel") and view.GenLevel:
                lvlid = view.GenLevel.Id
                elems = [e for e in elems if getattr(e, "LevelId", None) == lvlid]
        except:
            pass
    return elems

def _select_elements(uidoc, elements):
    # IronPython-friendly: convert to List[ElementId]
    from System.Collections.Generic import List
    from Autodesk.Revit.DB import ElementId
    ids = List[ElementId]([e.Id for e in elements])
    uidoc.Selection.SetElementIds(ids)

def _clear_selection(uidoc):
    from System.Collections.Generic import List
    from Autodesk.Revit.DB import ElementId
    uidoc.Selection.SetElementIds(List[ElementId]())
# ------------------------------------------------------------------------
# Main agent logic
# ------------------------------------------------------------------------
def run_agent():
    cfg = load_config()
    ctx = build_ctx()
    stop_on_fail = bool(cfg.get("stop_on_fail", True))
    dry_run = bool(cfg.get("dry_run", False))
    ctx["dry_run"] = dry_run

    log("Starting {0} v{1}".format(cfg.get("agent_name", "Agent"), cfg.get("version", "0.0")))
    log("DRY_RUN={0} | STOP_ON_FAIL={1}".format(dry_run, stop_on_fail))

    plan = cfg.get("plan", [])
    if not plan:
        log("No steps in plan; nothing to do.")
        return

    results = []
    for i, step in enumerate(plan):
        tool_name = step.get("tool")
        log("Step {0}/{1}: executing '{2}'".format(i + 1, len(plan), tool_name))

        preselected = False
        try:
            # Preselect Spaces for place_receptacles (so the tool runs the same as when you click by hand)
            if tool_name == "place_receptacles":
                spaces = _collect_spaces(ctx["doc"], mode="all")  # or "active_view_level"
                log("[DIAG] Preselecting {0} Space(s) for place_receptacles.".format(len(spaces)))
                if spaces:
                    _select_elements(ctx["uidoc"], spaces)
                    preselected = True
                else:
                    log("[HINT] No MEP Spaces found; the tool may no-op.")

            execute_tool_script(tool_name, ctx)
            results.append({"tool": tool_name, "ok": True})

        except Exception as e:
            tb = traceback.format_exc()
            log("!! '{0}' failed: {1}\n{2}".format(tool_name, e, tb))
            results.append({"tool": tool_name, "ok": False, "error": str(e)})
            if stop_on_fail:
                log("STOP_ON_FAIL=True. Halting.")
                break
        finally:
            if tool_name == "place_receptacles" and preselected:
                _clear_selection(ctx["uidoc"])

    ok_ct   = sum(1 for r in results if r.get("ok"))
    fail_ct = sum(1 for r in results if not r.get("ok"))
    log("Agent complete. {0} succeeded, {1} failed.".format(ok_ct, fail_ct))
    info_dialog("Agent", "Done.\nSucceeded: {0}\nFailed: {1}".format(ok_ct, fail_ct))

if __name__ == "__main__":
    try:
        run_agent()
    except Exception as ex:
        tb = traceback.format_exc()
        log("FATAL: {0}\n{1}".format(ex, tb))
        info_dialog("Agent Error", str(ex))