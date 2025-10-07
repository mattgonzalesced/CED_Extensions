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
def execute_tool_script(tool_name, ctx, exec_ctx="project"):
    """Run tools/<tool_name>.py as if launched by a pyRevit pushbutton."""
    import os, sys, runpy
    # IronPython-compatible builtins
    try:
        import builtins as _builtins
    except ImportError:
        import __builtin__ as _builtins

    import pyrevit as _py

    # Resolve UIApplication first (so pyrevit.revit can use it)
    uiapp = getattr(_py, "__revit__", None)
    try:
        if uiapp is None and hasattr(_py, "HOST_APP") and getattr(_py.HOST_APP, "uiapp", None):
            uiapp = _py.HOST_APP.uiapp
    except Exception:
        pass
    if uiapp is None and hasattr(_builtins, "__revit__"):
        uiapp = getattr(_builtins, "__revit__", None)
    if uiapp is None:
        raise EnvironmentError("Can’t resolve UIApplication. Run from pyRevit with a project open.")

    # Make pyrevit’s package-level global available BEFORE importing pyrevit.revit
    setattr(_py, "__revit__", uiapp)

    from pyrevit import revit as _revit
    from pyrevit import script as _script
    from pyrevit import forms as _forms
    try:
        from pyrevit import HOST_APP as _HOST_APP
    except Exception:
        _HOST_APP = None

    uidoc = getattr(_revit, "uidoc", None) or getattr(uiapp, "ActiveUIDocument", None)
    doc   = getattr(_revit, "doc", None)   or (getattr(uidoc, "Document", None) if uidoc else None)
    if doc is None or (hasattr(doc, "IsFamilyDocument") and doc.IsFamilyDocument):
        raise EnvironmentError("Open a project document (not a family) and run from a pyRevit button.")

    tool_path = os.path.join(TOOLS_DIR, tool_name + ".py")
    if not os.path.exists(tool_path):
        raise RuntimeError("Tool not found: {0}".format(tool_path))

    # >>> Inject pyRevit-like globals (including the exec context) <<<
    globals_dict = {
        "__name__": "__main__",
        "__file__": tool_path,
        "__revit__": uiapp,          # UIApplication
        "__uidoc__": uidoc,
        "__doc__": doc,
        "__context__": exec_ctx,     # <<< "project" or "selection"
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
    setattr(_builtins, "__context__", exec_ctx)
    setattr(_builtins, "revit", _revit)
    setattr(_builtins, "script", _script)
    setattr(_builtins, "forms", _forms)
    setattr(_builtins, "HOST_APP", _HOST_APP)

    # Env hints some scripts read
    os.environ["PYREVIT_RUNNING"] = "1"
    os.environ["PYREVIT_EXEC_CTX"] = exec_ctx

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
        # cleanup
        for k in ("__revit__", "__uidoc__", "__doc__", "__context__", "revit", "script", "forms", "HOST_APP"):
            if hasattr(_builtins, k):
                try: delattr(_builtins, k)
                except: pass

# --- Agent-side selection helpers (no tool edits needed) ---
# --- Selection helpers (IronPython-friendly) ---
def _save_selection(uidoc):
    # returns a plain Python list of ElementId
    return list(uidoc.Selection.GetElementIds())

def _restore_selection(uidoc, saved_ids):
    from System.Collections.Generic import List
    from Autodesk.Revit.DB import ElementId
    uidoc.Selection.SetElementIds(List[ElementId](saved_ids))

def _set_selection_elements(uidoc, elements):
    from System.Collections.Generic import List
    from Autodesk.Revit.DB import ElementId
    uidoc.Selection.SetElementIds(List[ElementId]([e.Id for e in elements]))

def _collect_spaces(doc, scope="all"):
    from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
    elems = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_MEPSpaces
    ).WhereElementIsNotElementType().ToElements()
    if scope == "active_view_level":
        try:
            view = doc.ActiveView
            genlvl = getattr(view, "GenLevel", None)
            if genlvl:
                lvlid = genlvl.Id
                elems = [e for e in elems if getattr(e, "LevelId", None) == lvlid]
        except:
            pass
    return elems

# Optional: default selection policy per tool (extend as you add tools)
SELECTION_POLICIES = {
    # run with Spaces preselected so the tool takes the same branch as manual use
    "place_receptacles": {"type": "spaces", "scope": "all"},   # or "active_view_level"
    # "some_other_tool": {"type": "rooms", "scope": "active_view_level"},
}


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
        # step may be {"tool":"name"} or richer; normalize
        tool_name = step.get("tool") if isinstance(step, dict) else str(step)
        log("Step {0}/{1}: executing '{2}'".format(i + 1, len(plan), tool_name))

        # --- selection push ---
        saved_ids = _save_selection(ctx["uidoc"])
        exec_ctx = "project"  # default

        # Decide selection based on per-tool policy (or per-step override if you add one)
        policy = SELECTION_POLICIES.get(tool_name)
        try:
            if policy:
                if policy.get("type") == "spaces":
                    spaces = _collect_spaces(ctx["doc"], scope=policy.get("scope", "all"))
                    log("[DIAG] Preselecting {0} Space(s) for {1}.".format(len(spaces), tool_name))
                    if spaces:
                        _set_selection_elements(ctx["uidoc"], spaces)
                        exec_ctx = "selection"
                    else:
                        log("[HINT] No MEP Spaces found; {0} may no-op.".format(tool_name))

            # run the tool with the appropriate context
            execute_tool_script(tool_name, ctx, exec_ctx=exec_ctx)
            results.append({"tool": tool_name, "ok": True})

        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            log("!! '{0}' failed: {1}\n{2}".format(tool_name, e, tb))
            results.append({"tool": tool_name, "ok": False, "error": str(e)})
            if stop_on_fail:
                log("STOP_ON_FAIL=True. Halting.")
                break
        finally:
            # --- selection pop (always restore) ---
            _restore_selection(ctx["uidoc"], saved_ids)

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