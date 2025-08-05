# -*- coding: utf-8 -*-
__title__ = "Parameter Linker"

from pyrevit import script, revit, DB
doc    = revit.doc
output = script.get_output()

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

# —— Output —— #
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