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
    """
    Execute a tools/<tool_name>.py as if launched by a pyRevit pushbutton.
    """
    import os, sys, runpy, builtins
    from pyrevit import revit as _revit
    from pyrevit import script as _script
    from pyrevit import forms as _forms
    from pyrevit import HOST_APP as _HOST_APP

    tool_path = os.path.join(TOOLS_DIR, tool_name + ".py")
    if not os.path.exists(tool_path):
        raise RuntimeError("Tool not found: {0}".format(tool_path))

    # Make the tool's module namespace look like a pyRevit pushbutton run
    globals_dict = {
        "__name__": "__main__",         # run as a script
        "__file__": tool_path,          # scripts often resolve paths from this
        "__revit__": _revit,            # legacy convenience
        "__doc__": _revit.doc,          # legacy convenience (module-level alias)
        "__uidoc__": _revit.uidoc,      # legacy convenience (module-level alias)
        # Common pyRevit imports some scripts expect to already exist
        "revit": _revit,
        "script": _script,
        "forms": _forms,
        "HOST_APP": _HOST_APP,
        # Optional agent context if tools want it
        "CTX": ctx,
    }

    # Also mirror these into builtins for very old scripts that read them there
    builtins.__revit__ = _revit
    builtins.revit = _revit
    builtins.script = _script
    builtins.forms = _forms
    builtins.HOST_APP = _HOST_APP
    # (do NOT overwrite builtins.__doc__ — that's Python's docstring)

    old_cwd = os.getcwd()
    old_sys_path0 = sys.path[0] if sys.path else None
    try:
        # Match a normal button's working directory
        os.chdir(os.path.dirname(tool_path))
        # Ensure the tool’s folder is at sys.path[0] like a direct run
        sys.path.insert(0, os.path.dirname(tool_path))

        log("Running tool: {0}".format(tool_name))
        log("CWD={0} | FILE={1}".format(os.getcwd(), tool_path))
        runpy.run_path(tool_path, globals_dict)
        log("Completed: {0}".format(tool_name))
    finally:
        os.chdir(old_cwd)
        # Remove the path we inserted at position 0
        if sys.path and sys.path[0] == os.path.dirname(tool_path):
            sys.path.pop(0)
        # Best-effort cleanup of builtins (harmless if left, but tidy)
        for k in ("__revit__", "revit", "script", "forms", "HOST_APP"):
            if hasattr(builtins, k):
                try: delattr(builtins, k)
                except: pass
# ------------------------------------------------------------------------
# Main agent logic
# ------------------------------------------------------------------------
def run_agent():
    cfg = load_config()
    ctx = build_ctx()
    stop_on_fail = bool(cfg.get("stop_on_fail", True))
    dry_run = bool(cfg.get("dry_run", False))

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
        try:
            execute_tool_script(tool_name, ctx)
            results.append({"tool": tool_name, "ok": True})
        except Exception as e:
            tb = traceback.format_exc()
            log("!! '{0}' failed: {1}\n{2}".format(tool_name, e, tb))
            results.append({"tool": tool_name, "ok": False, "error": str(e)})
            if stop_on_fail:
                log("STOP_ON_FAIL=True. Halting.")
                break

    ok_ct   = len([r for r in results if r.get("ok")])
    fail_ct = len([r for r in results if not r.get("ok")])
    log("Agent complete. {0} succeeded, {1} failed.".format(ok_ct, fail_ct))
    info_dialog("Agent", "Done.\nSucceeded: {0}\nFailed: {1}".format(ok_ct, fail_ct))

if __name__ == "__main__":
    try:
        run_agent()
    except Exception as ex:
        tb = traceback.format_exc()
        log("FATAL: {0}\n{1}".format(ex, tb))
        info_dialog("Agent Error", str(ex))