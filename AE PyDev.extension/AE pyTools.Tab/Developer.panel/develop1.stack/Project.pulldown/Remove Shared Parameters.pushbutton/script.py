# -*- coding: utf-8 -*-

import os

from pyrevit import revit, forms, script

doc = revit.doc
app = doc.Application
logger = script.get_logger()
output = script.get_output()


def pick_shared_parameter_file():
    return forms.pick_file(file_ext="txt", title="Select Shared Parameter File")


def load_shared_param_file_definitions(filepath):
    """
    Returns a list of tuples: (name, guid_string_or_blank)
    Restores original SharedParametersFilename.
    """
    original_file = app.SharedParametersFilename
    try:
        app.SharedParametersFilename = filepath
        sp_file = app.OpenSharedParameterFile()
        if not sp_file:
            raise Exception("Failed to open shared parameter file.")

        defs = []
        for group in sp_file.Groups:
            for d in group.Definitions:
                guid_str = ""
                try:
                    guid_str = str(d.GUID)
                except:
                    guid_str = ""
                defs.append((d.Name, guid_str))
        return defs
    finally:
        app.SharedParametersFilename = original_file


def get_project_parameter_definitions():
    """
    Returns a list of tuples: (definition, name, guid_string_or_blank)
    Includes ALL project parameters in ParameterBindings.
    """
    defs = []
    bindings = doc.ParameterBindings
    it = bindings.ForwardIterator()
    it.Reset()

    while it.MoveNext():
        d = it.Key
        name = ""
        guid_str = ""
        try:
            name = d.Name
        except:
            name = "<unnamed>"

        # Only shared parameters have a GUID; for others this will fail
        try:
            guid_str = str(d.GUID)
        except:
            guid_str = ""

        defs.append((d, name, guid_str))
    return defs


def build_candidate_rows(file_defs, proj_defs):
    """
    Candidates are those where file name matches a project param name.
    Returns dict display_string -> project Definition
    Also returns audit rows for printing.
    """
    # index project params by name
    proj_by_name = {}
    for d, name, guid_str in proj_defs:
        if name not in proj_by_name:
            proj_by_name[name] = []
        proj_by_name[name].append((d, guid_str))

    candidates = {}
    audit_rows = []

    for file_name, file_guid in file_defs:
        if file_name not in proj_by_name:
            continue

        # There can be multiple project params with same name (rare, but possible)
        for proj_def, proj_guid in proj_by_name[file_name]:
            status = ""
            if file_guid and proj_guid:
                if file_guid == proj_guid:
                    status = "GUID MATCH"
                else:
                    status = "NAME ONLY (GUID DIFFERENT)"
            else:
                status = "NAME ONLY (NO GUID)"

            display = "{0} | {1}".format(file_name, status)
            # keep display unique
            if display in candidates:
                display = "{0} | proj_guid={1}".format(display, proj_guid)

            candidates[display] = proj_def
            audit_rows.append((file_name, status, file_guid, proj_guid))

    return candidates, audit_rows


def print_audit(audit_rows):
    if not audit_rows:
        return

    output.print_md("### Match Audit (File vs Project)")
    output.print_md("| Parameter | Status | File GUID | Project GUID |")
    output.print_md("|---|---|---|---|")
    for name, status, fguid, pguid in sorted(audit_rows, key=lambda x: (x[1], x[0])):
        output.print_md("| {0} | {1} | {2} | {3} |".format(name, status, fguid or "", pguid or ""))


def main():
    shared_param_path = pick_shared_parameter_file()
    if not shared_param_path:
        return

    if not os.path.exists(shared_param_path):
        forms.alert("File does not exist.", title="Error")
        return

    file_defs = load_shared_param_file_definitions(shared_param_path)
    proj_defs = get_project_parameter_definitions()

    candidates, audit_rows = build_candidate_rows(file_defs, proj_defs)

    if not candidates:
        # This means even NAME matching found nothing, which would contradict your "it worked before"
        forms.alert(
            "No matches found.\n\n"
            "This tool matches by NAME first (like the original version), then audits GUID.\n"
            "If you expected matches, the picked shared parameter file may not contain those names.",
            title="No Matches"
        )
        return

    print_audit(audit_rows)

    selected = forms.SelectFromList.show(
        sorted(candidates.keys()),
        title="Select Project Parameters to Delete",
        multiselect=True
    )

    if not selected:
        return

    to_delete = [candidates[s] for s in selected]

    confirm = forms.alert(
        "Delete selected PROJECT PARAMETERS?\n\n"
        "{}\n\nThis cannot be undone.".format("\n".join([d.Name for d in to_delete])),
        title="Confirm Deletion",
        ok=False,
        yes=True,
        no=True
    )
    if not confirm:
        return

    removed = []
    failed = []

    with revit.Transaction("Delete Selected Project Parameters"):
        bindings = doc.ParameterBindings
        for d in to_delete:
            try:
                if bindings.Remove(d):
                    removed.append(d.Name)
                else:
                    failed.append(d.Name)
            except Exception as ex:
                logger.debug("Failed removing {}: {}".format(d.Name, ex))
                failed.append(d.Name)

    if removed:
        output.print_md("### ✅ Removed")
        for n in sorted(removed):
            output.print_md("- {}".format(n))

    if failed:
        output.print_md("### ⚠️ Failed")
        for n in sorted(failed):
            output.print_md("- {}".format(n))

    forms.alert("Removed: {}\nFailed: {}".format(len(removed), len(failed)), title="Done")


if __name__ == "__main__":
    main()
