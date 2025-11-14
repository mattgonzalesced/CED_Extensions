# -*- coding: utf-8 -*-
__title__ = "Parameter Linker"

import csv
import os

from pyrevit import script, revit, DB, forms

doc    = revit.doc
output = script.get_output()
logger = script.get_logger()
def _unwrap_parameter_id(rule):
    """Return (param_name, ElementId) or ('<unknown>', None)"""
    raw = rule.GetRuleParameter()
    if isinstance(raw, DB.ElementId):
        pid = raw
    else:
        try:
            pid = raw.ParameterId
        except:
            return ("<unknown>", None)
    name = "<unknown>"
    try:
        elem = doc.GetElement(pid)
        name = elem.Definition.Name
    except:
        name = str(pid.IntegerValue)
    return (name, pid.IntegerValue)

def _get_rule_value(rule):
    """Return the rule’s comparison value (string or numeric)."""
    if hasattr(rule, "RuleString"):
        return rule.RuleString
    if hasattr(rule, "RuleValue"):
        return rule.RuleValue
    if hasattr(rule, "GetIntegerValue"):
        return rule.GetIntegerValue()
    if hasattr(rule, "Value"):
        return rule.Value
    if hasattr(rule, "Threshold"):
        return rule.Threshold
    return "<no-value>"

def parse_filter_rules(filt, level=0):
    """Recursively flatten any AND/OR groups into per-rule lines."""
    indent = "  " * level
    lines = []

    # unwrap AND/OR groups entirely
    if isinstance(filt, (DB.LogicalAndFilter, DB.LogicalOrFilter)):
        for sub in filt.GetFilters():
            lines += parse_filter_rules(sub, level+1)
        return lines

    # leaf: ElementParameterFilter
    if isinstance(filt, DB.ElementParameterFilter):
        for rule in filt.GetRules():
            inverted = isinstance(rule, DB.FilterInverseRule)
            if inverted:
                rule = rule.GetInnerRule()

            pname, _ = _unwrap_parameter_id(rule)
            evaluator = rule.GetEvaluator().__class__.__name__
            if inverted:
                if evaluator.endswith("Contains"):
                    evaluator = evaluator.replace("Contains", "DoesNotContain")
                else:
                    evaluator = "Not" + evaluator

            val     = _get_rule_value(rule)
            red_val = '<span style="color:red">{0}</span>'.format(val)

            line = "{0}- **{1}**: *{2}* {3}".format(
                indent, pname, evaluator, red_val
            )
            lines.append(line)
        return lines

    # fallback for any other filter type
    try:
        tname = filt.GetType().Name
    except:
        tname = filt.__class__.__name__
    lines.append("{0}- **UnknownFilterType:{1}**".format(indent, tname))
    return lines

def document_filters():
    """Collect each ParameterFilterElement’s categories + parsed rules."""
    result = {}
    collector = DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement)
    for pf in collector.ToElements():
        fname = pf.Name or "<Unnamed>"
        cats  = []
        for cid in pf.GetCategories():
            c = DB.Category.GetCategory(doc, cid)
            cats.append(c.Name if c else "<none>")
        cats = sorted(cats)

        try:
            rules = parse_filter_rules(pf.GetElementFilter())
        except Exception as ex:
            rules = ["- ERROR: {0}".format(ex)]

        result[fname] = {
            "Categories": cats,
            "Rules":      rules
        }
    return result

def report_workset_filters():
    """Identify and report filters that reference the Workset parameter (-1002053)."""
    target_pid = DB.ElementId(-1002053)
    matching_filters = []

    def contains_workset_rule(filt):
        """Recursively check if a filter or its children use the Workset parameter."""
        # Handle logical filters
        if isinstance(filt, (DB.LogicalAndFilter, DB.LogicalOrFilter)):
            for sub in filt.GetFilters():
                if contains_workset_rule(sub):
                    return True
            return False

        # Handle element parameter filters
        if isinstance(filt, DB.ElementParameterFilter):
            for rule in filt.GetRules():
                # Handle inverse rules
                if isinstance(rule, DB.FilterInverseRule):
                    rule = rule.GetInnerRule()

                try:
                    raw_param = rule.GetRuleParameter()
                    if isinstance(raw_param, DB.ElementId):
                        pid = raw_param
                    else:
                        pid = raw_param.ParameterId

                    if pid == target_pid:
                        return True
                except Exception:
                    continue
        return False

    # Collect all ParameterFilterElements
    collector = DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement)

    for pf in collector.ToElements():
        try:
            if contains_workset_rule(pf.GetElementFilter()):
                matching_filters.append(pf.Name)
        except Exception:
            continue

    matching_filters = sorted(set(matching_filters))
    total = len(matching_filters)

    output.print_md("# Filters Using Workset Parameter")
    if total == 0:
        output.print_md("_No filters reference the Workset parameter (-1002053)._")
        return

    for name in matching_filters:
        output.print_md("- **{0}**".format(name))

    output.print_md("\n**Total Filters:** {0}".format(total))

# —— Output —— #
def report_all_filters():
    filters = document_filters()

    output.print_md("# ParameterFilterElement Rules")
    for fname in sorted(filters):
        info = filters[fname]
        output.print_md("## Filter: {0}".format(fname))
        output.print_md("**Categories**: {0}".format(
            ", ".join(info["Categories"]) or "<none>"
        ))
        output.print_md("**Rules**:")
        for rule_line in info["Rules"]:
            output.print_md(rule_line)


# -------------------------------------------------------------
# FILTER RULES
# -------------------------------------------------------------
IGNORE_BICS = [
    DB.BuiltInCategory.OST_Cameras,
    DB.BuiltInCategory.OST_ProjectBasePoint,
    DB.BuiltInCategory.OST_SharedBasePoint,
    DB.BuiltInCategory.OST_SitePoint,
    DB.BuiltInCategory.OST_IOS_GeoLocations,
]
IGNORE_TYPES = (DB.RevitLinkInstance, DB.ImportInstance)

# -------------------------------------------------------------
# HELPERS
# -------------------------------------------------------------
def is_valid_element(el):
    """Filter out unwanted elements."""
    if not el or el.Id.IntegerValue < 0:
        return False
    cat = el.Category
    if cat:
        if cat.CategoryType == DB.CategoryType.Annotation:
            return False
        if cat.Id.IntegerValue in [int(bic) for bic in IGNORE_BICS]:
            return False
    if isinstance(el, IGNORE_TYPES):
        return False
    if isinstance(el, DB.View3D) and el.IsPerspective:
        return False
    return True


def get_visible_elements(view):
    """Return dict {ElementId: CategoryName} for valid visible elements."""
    fec = DB.FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType()
    result = {}
    for el in fec:
        if is_valid_element(el):
            cat = el.Category
            cat_name = cat.Name if cat else "<No Category>"
            result[el.Id.IntegerValue] = cat_name
    return result


def group_by_category(id_dict):
    """Return {Category: [ElementIds]}."""
    grouped = {}
    for eid, cat in id_dict.items():
        grouped.setdefault(cat, []).append(eid)
    return grouped


def print_category_summary(title, grouped_dict):
    output.print_md("#### {}".format(title))
    if not grouped_dict:
        output.print_md("_None_")
        return
    for cat, ids in sorted(grouped_dict.items()):
        output.print_md("- **{}:** {} elements".format(cat, len(ids)))


# -------------------------------------------------------------
# MAIN LOGIC
# -------------------------------------------------------------
def analyze_views(baseline_view, coord_views):
    baseline = get_visible_elements(baseline_view)
    baseline_ids = set(baseline.keys())

    seen = set()
    duplicates_dict = {}
    missing_ids = set(baseline_ids)

    for v in coord_views:
        vis = get_visible_elements(v)
        vis_ids = set(vis.keys())
        overlap = seen.intersection(vis_ids)
        if overlap:
            duplicates_dict[v.Name] = dict((i, vis[i]) for i in overlap if i in vis)
        seen.update(vis_ids)
        missing_ids -= vis_ids

    missing_dict = dict((i, baseline[i]) for i in missing_ids if i in baseline)
    return duplicates_dict, missing_dict


# -------------------------------------------------------------
# CSV EXPORT
# -------------------------------------------------------------
def export_to_csv(coord_views, duplicates, missing):
    """Export only element IDs to CSV (grouped by view)."""
    # Use pyRevit temp directory instead of get_file()
    folder = forms.pick_folder()
    csv_path = os.path.join(folder, "Coordination_View_Analysis.csv")

    with open(csv_path, "w", newline="") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["Coordination View", "Element ID"])

        # All view IDs
        for v in coord_views:
            writer.writerow([v.Name])
            for eid in sorted(get_visible_elements(v).keys()):
                writer.writerow(["", eid])

        if duplicates:
            writer.writerow([])
            writer.writerow(["Duplicate Elements Across Views"])
            for vname, id_dict in duplicates.items():
                writer.writerow([vname])
                for eid in sorted(id_dict.keys()):
                    writer.writerow(["", eid])

        if missing:
            writer.writerow([])
            writer.writerow(["Missing Elements (not shown anywhere)"])
            for eid in sorted(missing.keys()):
                writer.writerow(["", eid])

    output.print_md("📁 **Exported CSV:** [{}]({})".format(os.path.basename(csv_path), csv_path))
    return csv_path


# -------------------------------------------------------------
# RUN
# -------------------------------------------------------------
def main():
    views = DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements()
    view_dict = {v.Name: v for v in views if not v.IsTemplate and not v.Name.startswith("<")}

    baseline_name = forms.SelectFromList.show(sorted(view_dict.keys()), title="Select Baseline View")
    if not baseline_name:
        return
    baseline_view = view_dict[baseline_name]

    coord_names = forms.SelectFromList.show(sorted(view_dict.keys()), multiselect=True, title="Select Coordination Views")
    if not coord_names:
        return
    coord_views = [view_dict[n] for n in coord_names]

    output.close_others()
    output.print_md("## Coordination Coverage Report")
    output.print_md("**Baseline:** `{}`".format(baseline_view.Name))
    output.print_md("**Coordination Views:** {}".format(", ".join(coord_names)))
    output.print_md("---")

    duplicates, missing = analyze_views(baseline_view, coord_views)

    # Per-view totals
    for v in coord_views:
        count = len(get_visible_elements(v))
        output.print_md("- **{}:** {} elements".format(v.Name, count))
    output.print_md("---")

    # Duplicates section
    total_dups = sum(len(ids) for ids in duplicates.values())
    if total_dups:
        output.print_md("### ⚠️ Duplicates Found: {}".format(total_dups))
        for vname, id_dict in duplicates.items():
            grouped = group_by_category(id_dict)
            output.print_md("##### {}".format(vname))
            print_category_summary("By Category", grouped)
    else:
        output.print_md("✅ No duplicates across coordination views.")

    # Missing section
    if missing:
        output.print_md("---")
        output.print_md("### ⚠️ Missing from All Coordination Views: {}".format(len(missing)))
        grouped_missing = group_by_category(missing)
        print_category_summary("By Category", grouped_missing)
    else:
        output.print_md("✅ All baseline elements accounted for.")

    export_to_csv(coord_views, duplicates, missing)


if __name__ == "__main__":
    main()



# # —— Output: parameters by category —— #
# params_by_cat = document_params_by_category()
# output.print_md("\n# Parameters by Category\n")
# for fname in sorted(params_by_cat):
#     output.print_md("## Filter: {0}".format(fname))
#     for cat in sorted(params_by_cat[fname].keys()):
#         names   = params_by_cat[fname][cat]
#         joined  = ", ".join(names) or "<none>"
#         output.print_md("- **{0}**: {1}".format(cat, joined))
#
#












# class ParameterSet:
#     def __init__(self):
#         # A list of parameter mappings (pairs of parameters to compare)
#         self.parameter_mappings = []  # List of tuples [(param_a, param_b), ...]
#
#     def add_mapping(self, param_a_metadata, param_b_metadata):
#         """Adds a parameter mapping to the set."""
#         self.parameter_mappings.append((param_a_metadata, param_b_metadata))
#
#     def get_mappings(self):
#         """Returns all parameter mappings."""
#         return self.parameter_mappings
#
# class ParameterMetadata:
#     def __init__(self, name, guid, param_id, storage_type, is_read_only, built_in_param=None):
#         self.name = name  # Parameter name
#         self.guid = guid  # GUID for shared parameters
#         self.param_id = param_id  # Revit Parameter ID
#         self.storage_type = storage_type  # StorageType (e.g., Integer, String, etc.)
#         self.is_read_only = is_read_only  # Whether the parameter is read-only
#         self.built_in_param = built_in_param  # BuiltInParameter enum (if applicable)
#
#
#     def to_dict(self):
#         """Returns a dictionary representation of the metadata."""
#         return {
#             "name": self.name,
#             "guid": str(self.guid) if self.guid else None,
#             "param_id": self.param_id,
#             "storage_type": self.storage_type,
#             "is_read_only": self.is_read_only,
#             "built_in_param": self.built_in_param,
#         }
#
#
#
#
# # Define metadata for parameters
# param_a1 = ParameterMetadata(
#     name="FLA Input_CED",
#     guid="54564ea7-fc79-44f8-9beb-c9b589901dee",
#     param_id=23926625,
#     storage_type="Double",
#     is_read_only=False,
#     built_in_param=False
# )
#
# param_a2 = ParameterMetadata(
#     name="Voltage_CED",
#     guid="04342884-6218-495e-970a-1cdd49f5ddc0",
#     param_id=23926634,
#     storage_type="Double",
#     is_read_only=False,
#     built_in_param=False
# )
#
# param_a3 = ParameterMetadata(
#     name="Phase_CED",
#     guid="d4252307-22ba-4917-b756-f79be1334c48",
#     param_id=23926632,
#     storage_type="Integer",
#     is_read_only=True,
#     built_in_param=False
# )
#
# param_b1 = ParameterMetadata(
#     name="CED-E-FLA",
#     guid=None,
#     param_id=2001,
#     storage_type="Double",
#     is_read_only=False,
#     built_in_param=False
# )
#
# param_b2 = ParameterMetadata(
#     name="VOLTAGE",
#     guid=None,
#     param_id=2002,
#     storage_type="String",
#     is_read_only=False,
#     built_in_param=False
# )
#
# param_b3 = ParameterMetadata(
#     name="PHASE",
#     guid=None,
#     param_id=2003,
#     storage_type="Integer",
#     is_read_only=False,
#     built_in_param=False
# )
#
# # Create a ParameterSet and add mappings
# parameter_set = ParameterSet()
# parameter_set.add_mapping(param_a1, param_b1)
# parameter_set.add_mapping(param_a2, param_b2)
# parameter_set.add_mapping(param_a3, param_b3)
#
# # Retrieve and display mappings
# for param_a, param_b in parameter_set.get_mappings():
#     print("Mapping:")
#     print("  Element A - Parameter:", param_a.to_dict())
#     print("  Element B - Parameter:", param_b.to_dict())
#
# for params_a, params_b in parameter_set.get_mappings():
#     params_a.to_dict()