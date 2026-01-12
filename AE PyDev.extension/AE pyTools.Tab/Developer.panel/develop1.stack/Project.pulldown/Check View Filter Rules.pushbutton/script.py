# -*- coding: utf-8 -*-

from pyrevit import revit, DB, script

doc = revit.doc
output = script.get_output()
output.close_others()

# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------
WORKSET_PARAM_ID = DB.ElementId(-1002053)
ONLY_SHOW_WORKSET_FILTERS = True  # <-- toggle this


# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------
def get_param_name_from_rule(rule):
    try:
        raw = rule.GetRuleParameter()
        if isinstance(raw, DB.ElementId):
            pid = raw
        else:
            pid = raw.ParameterId

        elem = doc.GetElement(pid)
        if elem and elem.Definition:
            return elem.Definition.Name
        return str(pid.IntegerValue)
    except Exception:
        return "<unknown>"


def get_rule_value(rule):
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


def evaluator_name(rule, inverted):
    name = rule.GetEvaluator().__class__.__name__
    if inverted:
        if "Contains" in name:
            return name.replace("Contains", "DoesNotContain")
        return "Not" + name
    return name


# ------------------------------------------------------------
# RULE PARSING (RECURSIVE, FLATTENED)
# ------------------------------------------------------------
def parse_filter(filt, rules_out, group=None):
    if isinstance(filt, DB.LogicalAndFilter):
        for sub in filt.GetFilters():
            parse_filter(sub, rules_out, "AND")
        return

    if isinstance(filt, DB.LogicalOrFilter):
        for sub in filt.GetFilters():
            parse_filter(sub, rules_out, "OR")
        return

    if isinstance(filt, DB.ElementParameterFilter):
        for rule in filt.GetRules():
            inverted = isinstance(rule, DB.FilterInverseRule)
            if inverted:
                rule = rule.GetInnerRule()

            pname = get_param_name_from_rule(rule)
            eval_name = evaluator_name(rule, inverted)
            val = get_rule_value(rule)

            rules_out.append({
                "group": group,
                "param": pname,
                "evaluator": eval_name,
                "value": val,
                "param_id": rule.GetRuleParameter()
            })
        return


# ------------------------------------------------------------
# MAIN COLLECTION
# ------------------------------------------------------------
def collect_filters():
    filters = []

    for pf in DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement):
        cats = []
        for cid in pf.GetCategories():
            c = DB.Category.GetCategory(doc, cid)
            if c:
                cats.append(c.Name)

        rules = []
        try:
            parse_filter(pf.GetElementFilter(), rules)
        except Exception as ex:
            rules.append({
                "group": None,
                "param": "<error>",
                "evaluator": "ERROR",
                "value": str(ex),
                "param_id": None
            })

        filters.append({
            "name": pf.Name,
            "categories": sorted(cats),
            "rules": rules
        })

    return sorted(filters, key=lambda x: x["name"].lower())


# ------------------------------------------------------------
# WORKSET FILTER CHECK
# ------------------------------------------------------------
def uses_workset_param(filter_data):
    for r in filter_data["rules"]:
        try:
            pid = r["param_id"]
            if isinstance(pid, DB.ElementId) and pid == WORKSET_PARAM_ID:
                return True
        except Exception:
            continue
    return False


# ------------------------------------------------------------
# OUTPUT
# ------------------------------------------------------------
def main():
    output.print_md("# Parameter Filter Review")

    filters = collect_filters()
    shown = 0

    for f in filters:
        if ONLY_SHOW_WORKSET_FILTERS and not uses_workset_param(f):
            continue

        shown += 1
        output.print_md("## {0}".format(f["name"]))
        output.print_md("**Categories:** {0}".format(", ".join(f["categories"]) or "<none>"))
        output.print_md("**Rules:**")

        for r in f["rules"]:
            line = "- **{0}**: *{1}* <span style=\"color:red\">{2}</span>".format(
                r["param"],
                r["evaluator"],
                r["value"]
            )
            if r["group"]:
                line = "{0} ({1})".format(line, r["group"])
            output.print_md(line)

    output.print_md("\n**Total Filters Shown:** {0}".format(shown))


# -*- coding: utf-8 -*-

from pyrevit import revit, DB, UI

doc = revit.doc


def purge_unused_material_assets(doc):
    # --------------------------------------------------
    # Collect materials and used asset ids
    # --------------------------------------------------
    used_appearance = set()
    used_structural = set()
    used_thermal = set()

    materials = DB.FilteredElementCollector(doc) \
        .OfClass(DB.Material) \
        .ToElements()

    for mat in materials:
        if mat.AppearanceAssetId != DB.ElementId.InvalidElementId:
            used_appearance.add(mat.AppearanceAssetId)

        if mat.StructuralAssetId != DB.ElementId.InvalidElementId:
            used_structural.add(mat.StructuralAssetId)

        if mat.ThermalAssetId != DB.ElementId.InvalidElementId:
            used_thermal.add(mat.ThermalAssetId)

    deleted_count = 0
    failed_count = 0

    # --------------------------------------------------
    # Transaction
    # --------------------------------------------------
    with revit.Transaction("Purge Unused Material Assets"):
        # Appearance assets
        appearance_assets = DB.FilteredElementCollector(doc) \
            .OfClass(DB.AppearanceAssetElement) \
            .ToElements()

        for asset in appearance_assets:
            if asset.Id not in used_appearance:
                try:
                    doc.Delete(asset.Id)
                    deleted_count += 1
                except Exception:
                    failed_count += 1

        # Structural + Thermal assets
        property_assets = DB.FilteredElementCollector(doc) \
            .OfClass(DB.PropertySetElement) \
            .ToElements()

        for asset in property_assets:
            if asset.Id not in used_structural and asset.Id not in used_thermal:
                try:
                    doc.Delete(asset.Id)
                    deleted_count += 1
                except Exception:
                    failed_count += 1

    # --------------------------------------------------
    # Summary only
    # --------------------------------------------------
    UI.TaskDialog.Show(
        "Material Asset Purge Complete",
        "Deleted assets: {}\nFailed deletions: {}".format(
            deleted_count, failed_count
        )
    )


# Run
purge_unused_material_assets(doc)


