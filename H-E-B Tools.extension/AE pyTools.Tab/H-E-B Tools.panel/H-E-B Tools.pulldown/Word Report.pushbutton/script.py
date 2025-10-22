# -*- coding: utf-8 -*-
###############################################################################
#  revision_report_console.py  (IronPython 2.7 - runs inside pyRevit)        #
###############################################################################

"""
Keeps the original console report **and** spawns a Python-3.13 subprocess that
builds a Word document.

• What you MUST edit once:
    PY3_EXE  – full path to your Python-3.13.1 interpreter
               (or just "python" if it’s on PATH)
• What you MAY edit:
    WORD_SCRIPT – location of revision_report_word.py if you store it elsewhere
"""

from __future__ import unicode_literals, print_function
from pyrevit import script, forms, coreutils, revit, DB
from pyrevit.revit import query
from pyrevit.compat import get_elementid_value_func

import os, sys, json, tempfile, subprocess
import datetime  # keep with the other imports near the top

###############################################################################
# --- CONFIG ------------------------------------------------------------------
###############################################################################
PY3_EXE      = r"python"          # EDIT if needed
WORD_SCRIPT  = os.path.join(os.path.dirname(__file__),
                             "revision_report_word.py")
get_id_value = get_elementid_value_func()

###############################################################################
# --- ORIGINAL HELPERS (unchanged logic) --------------------------------------
###############################################################################
console = script.get_output()
console.close_others(); console.set_height(800)
script_dir = os.path.dirname(__file__)
logo_path  = os.path.join(script_dir, 'CED_Logo_H.png')
doc        = revit.doc

def validate_additional_parameters(param_names):
    sample = (DB.FilteredElementCollector(doc)
              .OfCategory(DB.BuiltInCategory.OST_RevisionClouds)
              .WhereElementIsNotElementType()
              .FirstElement())
    return [] if not sample else [p for p in param_names
                                  if not sample.LookupParameter(p)]

def clean_param_name(p):
    for suf in ("_CED", "_CEDT"):
        if p.endswith(suf): return p[:-len(suf)]
    return p

def get_param_value_by_name(el, name):
    param = el.LookupParameter(name)
    val   = query.get_param_value(param) if param else None
    return val if val is not None else u""

def get_revision_data_by_sheet(param_names):
    data, sheets = {}, DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet)
    for sheet in sheets:
        sn = query.get_param_value(sheet.LookupParameter("Sheet Number"))
        for cid in sheet.GetAllRevisionCloudIds():
            cloud = doc.GetElement(cid)
            rid   = get_id_value(cloud.RevisionId)
            data.setdefault(rid, [])
            cmnt = query.get_param_value(cloud.LookupParameter("Comments"))
            if not cmnt: continue
            row = {"Sheet Number": sn, "Comments": cmnt}
            for p in param_names:
                row[p] = get_param_value_by_name(cloud, p)
            data[rid].append(row)
    return data

def deduplicate(items):
    seen, out = set(), []
    for it in items:
        key = (it.get("Sheet Number", u""),
               it.get("Comments", u"").strip())
        if key[1] and key not in seen:
            out.append(it); seen.add(key)
    return out

###############################################################################
# --- CONSOLE PREVIEW (unchanged tables) --------------------------------------
###############################################################################
def print_project_metadata(info):
    date = coreutils.current_date()
    console.print_html("<img src='{0}' width='150px' />".format(logo_path))
    console.print_md("**Coolsys Energy Design**")
    console.print_md("## Project Revision Summary\n---")
    console.print_md("Project Number: **{0}**".format(info.number))
    console.print_md("Client: **{0}**".format(info.client_name))
    console.print_md("Project Name: **{0}**".format(info.name))
    console.print_md("Report Date: **{0}**\n---".format(date))

def print_revision_report(revisions, data, param_names):
    base_cols           = ["Sheet Number", "Comments"]
    map_clean           = {clean_param_name(p): p for p in param_names}

    for rev in revisions:
        rid   = get_id_value(rev.Id)
        rn    = query.get_param_value(rev.LookupParameter("Revision Number"))
        rdt   = query.get_param_value(rev.LookupParameter("Revision Date"))
        rdesc = query.get_param_value(rev.LookupParameter("Revision Description"))
        console.print_md("### Revision Number: {0} | Date: {1} | Description: {2}"
                         .format(rn, rdt, rdesc))

        items = deduplicate(sorted(data.get(rid, []),
                                   key=lambda x: x.get("Sheet Number", u"")))
        if not items:
            console.print_md("No revision clouds found."); console.insert_divider(); continue

        cols = list(base_cols)
        for p in param_names:
            cp = clean_param_name(p)
            if cp not in cols: cols.append(cp)
        display = ["Description of Change" if c == "Comments" else c for c in cols]

        table = []
        for it in items:
            row = []
            for col in cols:
                if col in ("Sheet Number", "Comments"):
                    row.append(it.get(col, u""))
                else:
                    row.append(it.get(map_clean.get(col, col), u""))
            table.append(row)
        console.print_table(table, columns=display); console.insert_divider()

###############################################################################
# --- MAIN --------------------------------------------------------------------
###############################################################################
def main():
    cfg     = script.get_config("revision_parameters_config")
    params  = (getattr(cfg, "selected_param_names", u"") or u"").split(",") if getattr(cfg, "selected_param_names", u"") else []
    bad     = validate_additional_parameters(params)
    if bad:
        forms.alert("These parameters are missing on Revision Clouds:\n{0}\n\n"
                    "Using default output.".format("\n".join(bad)))
        params = []; cfg.selected_param_names = ""; script.save_config()

    revs = forms.select_revisions(button_name="Select Revision", multiple=True)
    if not revs: script.exit()

    data   = get_revision_data_by_sheet(params)
    info   = revit.query.get_project_info()
    # Console preview (unchanged)
    print_project_metadata(info); print_revision_report(revs, data, params)

    # ------------------------------------------------------------------ Docx part
    timestamp    = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    default_name = "Revision_Report_{}.docx".format(timestamp)

    save_path = forms.save_file(file_ext='docx',
                                title="Save Revision Report As")

    if not save_path: forms.alert("Cancelled."); script.exit()

    # Build JSON payload for script-2
    payload = {"project": {"number"      : info.number,
                           "client"      : info.client_name,
                           "name"        : info.name,
                           "report_date" : coreutils.current_date(),
                           "logo_path"   : logo_path},
               "columns_info": {},
               "revisions": []}

    base_cols           = ["Sheet Number", "Comments"]
    map_clean           = {clean_param_name(p): p for p in params}

    for rev in revs:
        rid   = get_id_value(rev.Id)
        rn    = query.get_param_value(rev.LookupParameter("Revision Number"))
        rdt   = query.get_param_value(rev.LookupParameter("Revision Date"))
        rdesc = query.get_param_value(rev.LookupParameter("Revision Description"))
        items = deduplicate(sorted(data.get(rid, []),
                                   key=lambda x: x.get("Sheet Number", u"")))
        cols = list(base_cols)
        for p in params:
            cp = clean_param_name(p)
            if cp not in cols: cols.append(cp)
        display = ["Description of Change" if c == "Comments" else c for c in cols]

        rows = []
        for it in items:
            row = []
            for col in cols:
                if col in ("Sheet Number", "Comments"):
                    row.append(it.get(col, u""))
                else:
                    row.append(it.get(map_clean.get(col, col), u""))
            rows.append(row)

        payload["revisions"].append(
            {"header" : "Revision Number: {0} | Date: {1} | Description: {2}"
                         .format(rn, rdt, rdesc),
             "columns": display,
             "rows"   : rows})

    # temp JSON file
    tmp_json = tempfile.NamedTemporaryFile(delete=False, suffix='.json')
    json.dump(payload, tmp_json, indent=2); tmp_json.close()

    # spawn subprocess
    try:
        WORD_EXE = os.path.join(os.path.dirname(__file__), "revision_report_word.exe")
        rc = subprocess.call([WORD_EXE, tmp_json.name, save_path])

        if rc == 0:
            forms.alert("Word report saved to:\n{0}".format(save_path))
        else:
            forms.alert("Python-3 script returned exit-code {0}".format(rc))
    finally:
        try: os.remove(tmp_json.name)
        except Exception: pass

if __name__ == "__main__":
    main()
