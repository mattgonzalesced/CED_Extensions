# -*- coding: utf-8 -*-
"""
Force a clean reload of every MEPRFP 2.0 lib module on script entry.

pyRevit's CPython engine keeps Python modules loaded in ``sys.modules``
across script runs in the same Revit session. That makes iterative
development painful — edits to a lib module aren't seen until Revit
restarts. Each pushbutton script calls ``purge()`` at the top of its
imports to drop our cached lib modules so the next ``import`` reads
the on-disk file fresh.

In production this costs one ``sys.modules`` scan per click (sub-ms,
ignorable). It does NOT touch pyRevit, pythonnet, vendored PyYAML,
.NET assemblies, or anything outside our lib.
"""

import sys


_LIB_MODULE_NAMES = frozenset({
    # data + storage
    "active_yaml",
    "schema",
    "schema_migrations",
    "yaml_io",
    "storage",
    "profile_model",
    "truth_groups",
    "element_linker",
    "element_linker_io",
    # capture / authoring
    "append_workflow",
    "capture",
    "directives",
    "directives_dialog",
    # lifecycle
    "merge_workflow",
    # editor (Manage Profiles depends on alias data shape)
    "manage_profiles_window",
    # placement
    "placement",
    "placement_window",
    "annotation_placement",
    "annotation_placement_window",
    "geometry",
    "hosted_annotations",
    "links",
    "selection",
    "shared_params",
    # audit
    "sync_audit",
    "sync_audit_window",
    # misc ops (Stage 4)
    "follow_parent_workflow",
    "follow_parent_window",
    "hide_profiles_workflow",
    "hide_profiles_window",
    "update_vector_workflow",
    "update_vector_window",
    "optimize_workflow",
    "optimize_window",
    "qaqc_workflow",
    "qaqc_window",
    # ui infra
    "forms_compat",
    "wpf",
    "wpf_dialogs",
})


def purge():
    """Drop cached MEPRFP 2.0 lib modules from ``sys.modules``.

    Safe to call repeatedly. Does not touch ``_dev_reload`` itself.
    """
    for name in list(sys.modules):
        head = name.split(".", 1)[0]
        if head in _LIB_MODULE_NAMES:
            del sys.modules[name]
