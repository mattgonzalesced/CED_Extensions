# -*- coding: utf-8 -*-
"""Run all configured proximity/clearance checks and show individual reports plus summary."""

__title__ = "Run All\nProximity Checks"
__doc__ = "Run all configured proximity-style quality checks and report each plus a summary."

import os
import sys

from pyrevit import forms, revit, script

# From AI_QC.pulldown/X.pushbutton: 5 levels up to extension root, then CEDLib.lib
def _lib_root():
    here = os.path.dirname(os.path.abspath(__file__))
    root = os.path.normpath(os.path.join(here, "..", "..", "..", "..", ".."))
    return os.path.join(root, "CEDLib.lib")

def _append_lib():
    lib = _lib_root()
    if lib not in sys.path:
        sys.path.insert(0, lib)

def _load_checks():
    _append_lib()
    from QualityChecks.quality_check_core import report_proximity_hits, summarize_results

    checks = []

    # Lights vs coils (panel-level module)
    try:
        import imp
        here = os.path.dirname(os.path.abspath(__file__))
        panel = os.path.normpath(os.path.join(here, "..", ".."))
        prox_path = os.path.join(panel, "Proximity Check.pushbutton", "proximity_lights_coils.py")
        if os.path.isfile(prox_path):
            mod = imp.load_source("ced_prox_lights_coils", prox_path)
            if hasattr(mod, "collect_hits"):
                checks.append({
                    "name": "Lights vs CED-R-KRACK coils",
                    "collector": lambda d: mod.collect_hits(d),
                    "title": "Proximity Check: Lights-Coils",
                    "subtitle": "Lighting fixtures within 18 inches of CED-R-KRACK coils",
                    "columns": ["Lighting ID", "Lighting Family : Type", "Coil ID", "Coil Family : Type", "Distance (in)"],
                })
    except Exception:
        pass

    # Library checks (proximity-style: a_id, a_label, b_id, b_label, distance_ft)
    for mod_name, cfg in [
        ("proximity_lights_sprinklers", {
            "name": "Lights vs Sprinklers",
            "title": "Proximity Check: Lights-Sprinklers",
            "subtitle": "Lighting fixtures within 18 inches of sprinklers",
            "columns": ["Lighting ID", "Lighting Family : Type", "Sprinkler ID", "Sprinkler Family : Type", "Distance (in)"],
            "keys": ("light_id", "light_label", "sprinkler_id", "sprinkler_label"),
        }),
        ("clearance_sprinkler_obstructions", {
            "name": "Sprinkler Obstruction Clearance",
            "title": "Sprinkler Obstruction Clearance",
            "subtitle": "Sprinklers within 18 inches of beams, duct, lights, or tray",
            "columns": ["Sprinkler ID", "Sprinkler Family : Type", "Obstruction ID", "Obstruction Family : Type", "Distance (in)"],
            "keys": ("sprinkler_id", "sprinkler_label", "obstruction_id", "obstruction_label"),
        }),
        ("clearance_equipment_access", {
            "name": "Equipment Access Clearance",
            "title": "Equipment Access Clearance",
            "subtitle": "Equipment with obstructions within 36 inches",
            "columns": ["Equipment ID", "Equipment Family : Type", "Obstruction ID", "Obstruction Family : Type", "Distance (in)"],
            "keys": ("equipment_id", "equipment_label", "obstruction_id", "obstruction_label"),
        }),
        ("mech_duct_clearance", {
            "name": "Duct Clearance",
            "title": "Duct Clearance (Mechanical)",
            "subtitle": "Duct vs structure (2\"), duct/pipe/tray (1.5\")",
            "columns": ["Duct ID", "Duct", "Other ID", "Other", "Distance (in)"],
            "keys": ("a_id", "a_label", "b_id", "b_label"),
        }),
        ("elec_conduit_pipe_clearance", {
            "name": "Conduit vs Pipe",
            "title": "Conduit vs Pipe Clearance (Electrical)",
            "subtitle": "Conduit vs hot 6\", vs cold/chilled 3\"",
            "columns": ["Conduit ID", "Conduit", "Pipe ID", "Pipe", "Distance (in)"],
            "keys": ("a_id", "a_label", "b_id", "b_label"),
        }),
        ("plumb_pipe_clearance", {
            "name": "Pipe Clearance (Plumbing)",
            "title": "Pipe Clearance (Plumbing)",
            "subtitle": "Pipe vs structure 2\", duct 1.5\", electrical 3\"",
            "columns": ["Pipe ID", "Pipe", "Other ID", "Other", "Distance (in)"],
            "keys": ("a_id", "a_label", "b_id", "b_label"),
        }),
        ("refrig_coils_proximity", {
            "name": "Refrigeration Coils Proximity",
            "title": "Refrigeration Coils Proximity",
            "subtitle": "Coils vs heat sources 24\", vs sprinklers 18\"",
            "columns": ["Coil ID", "Coil", "Other ID", "Other", "Distance (in)"],
            "keys": ("a_id", "a_label", "b_id", "b_label"),
        }),
    ]:
        try:
            mod = __import__("QualityChecks." + mod_name, fromlist=["collect_hits"])
            if hasattr(mod, "collect_hits"):
                k = cfg["keys"]
                def _collect(d, m=mod, o=cfg):
                    return m.collect_hits(d, options=None)
                checks.append({
                    "name": cfg["name"],
                    "collector": _collect,
                    "title": cfg["title"],
                    "subtitle": cfg["subtitle"],
                    "columns": cfg["columns"],
                    "keys": k,
                })
        except Exception:
            pass

    return report_proximity_hits, summarize_results, checks

def _map_hit(item, keys):
    a_id, a_lbl, b_id, b_lbl = keys
    return {
        "a_id": item.get(a_id),
        "a_label": item.get(a_lbl),
        "b_id": item.get(b_id),
        "b_label": item.get(b_lbl),
        "distance_ft": item.get("distance_ft"),
    }

def main():
    doc = getattr(revit, "doc", None)
    if doc is None or getattr(doc, "IsFamilyDocument", False):
        forms.alert("Open a project model before running quality checks.", title=__title__)
        return

    report_hits, summarize_results, checks = _load_checks()
    if not checks:
        forms.alert("No proximity quality checks are currently configured or available.", title=__title__)
        return

    output = script.get_output()
    output.set_width(1000)
    output.print_md("# Run All Proximity Checks")

    results = []
    for check in checks:
        name = check.get("name") or "<Unnamed>"
        collector = check.get("collector")
        if collector is None:
            continue
        try:
            raw = list(collector(doc))
        except Exception:
            raw = []

        keys = check.get("keys")
        if keys:
            mapped = [_map_hit(item, keys) for item in raw]
        else:
            mapped = [
                {
                    "a_id": item.get("light_id"),
                    "a_label": item.get("light_label"),
                    "b_id": item.get("coil_id") or item.get("sprinkler_id"),
                    "b_label": item.get("coil_label") or item.get("sprinkler_label"),
                    "distance_ft": item.get("distance_ft"),
                }
                for item in raw
            ]

        report_hits(
            title=check.get("title") or name,
            subtitle=check.get("subtitle") or "",
            hits=mapped,
            columns=check.get("columns"),
            show_empty=False,
        )
        results.append({"check_name": name, "hits": mapped, "pass": not bool(mapped)})

    summary = summarize_results(results)
    rows = [[r.get("check_name") or "<Unnamed>", "Pass" if r.get("pass") else "Fail", len(r.get("hits") or [])] for r in results]
    output.print_table(rows, columns=["Check", "Result", "Hit Count"])

    if summary.get("total_hits", 0) == 0:
        forms.alert("All proximity checks passed.\n\nNo issues were found.", title=__title__)
    else:
        forms.alert(
            "Completed {} check(s). {} reported issues ({} total hits).\n\nSee the output panel for details.".format(
                summary.get("total_checks", 0), summary.get("failing_checks", 0), summary.get("total_hits", 0)
            ),
            title=__title__,
        )

if __name__ == "__main__":
    main()
