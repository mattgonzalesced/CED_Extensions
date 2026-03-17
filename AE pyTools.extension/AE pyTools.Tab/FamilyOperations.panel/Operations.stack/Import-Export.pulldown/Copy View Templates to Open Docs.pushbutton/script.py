# -*- coding: utf-8 -*-
"""
Overwrite View Templates Across Open Models
- Reassigns views
- Reconciles filters safely
- Deletes old templates
- Renames new templates
"""

from System.Collections.Generic import List
from pyrevit import revit, DB, forms, script

logger = script.get_logger()
output = script.get_output()

src_doc = revit.doc


# ------------------------------------------------------------
# Helpers
# ------------------------------------------------------------

def get_view_templates_by_name(doc):
    templates = {}
    for v in DB.FilteredElementCollector(doc).OfClass(DB.View):
        if v.IsTemplate:
            templates[v.Name] = v
    return templates


def map_template_usage(doc):
    usage = {}
    for v in DB.FilteredElementCollector(doc).OfClass(DB.View):
        if not v.IsTemplate and v.ViewTemplateId != DB.ElementId.InvalidElementId:
            vt = doc.GetElement(v.ViewTemplateId)
            if vt:
                usage.setdefault(vt.Name, []).append(v)
    return usage


def get_all_filters_by_name(doc):
    filters = {}
    for f in DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement):
        filters[f.Name] = f
    return filters


def extract_filter_rules(element_filter):
    """Recursively extract FilterRule objects from an ElementFilter."""
    rules = []

    if isinstance(element_filter, DB.ElementParameterFilter):
        rules.extend(element_filter.GetRules())

    elif isinstance(element_filter, DB.LogicalAndFilter) or isinstance(element_filter, DB.LogicalOrFilter):
        for sub_filter in element_filter.GetFilters():
            rules.extend(extract_filter_rules(sub_filter))

    return rules


def filter_signature(filter_elem):
    # Categories
    cats = sorted([c.IntegerValue for c in filter_elem.GetCategories()])

    rules = []
    elem_filter = filter_elem.GetElementFilter()
    if not elem_filter:
        return (tuple(cats), tuple())

    for rule in extract_filter_rules(elem_filter):
        if isinstance(rule, DB.FilterInverseRule):
            rule = rule.GetInnerRule()

        try:
            param_id = rule.GetRuleParameter().IntegerValue
            evaluator = rule.GetEvaluator().GetType().FullName
            value = str(rule.GetValue())
        except Exception:
            continue

        rules.append((param_id, evaluator, value))

    return (tuple(cats), tuple(rules))



def update_filter_in_place(dest_filter, src_filter):
    logger.debug("    Updating filter rules: {}".format(dest_filter.Name))

    # Categories
    dest_filter.SetCategories(src_filter.GetCategories())

    # Replace element filter entirely
    dest_elem_filter = src_filter.GetElementFilter()
    if dest_elem_filter:
        dest_filter.SetElementFilter(dest_elem_filter)


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------

output.close_others()
output.print_md("## View Template Overwrite")

selected_templates = forms.select_viewtemplates(doc=src_doc)
if not selected_templates:
    forms.alert("No view templates selected.")
    script.exit()

dest_docs = forms.select_open_docs(title="Select Destination Documents")
if not dest_docs:
    forms.alert("No destination documents selected.")
    script.exit()

print("Selected {} templates".format(len(selected_templates)))
print("Selected {} destination docs".format(len(dest_docs)))

src_filters = get_all_filters_by_name(src_doc)

for dest_doc in dest_docs:
    print("======================================")
    print("Processing destination document")
    print(dest_doc.Title)

    with revit.Transaction("Overwrite View Templates", doc=dest_doc):

        # --- ANALYSIS ---
        print("Phase 1: Analyzing destination state")

        dest_templates = get_view_templates_by_name(dest_doc)
        template_usage = map_template_usage(dest_doc)
        dest_filters = get_all_filters_by_name(dest_doc)

        logger.debug("  Existing templates: {}".format(len(dest_templates)))
        logger.debug("  Existing filters: {}".format(len(dest_filters)))

        # --- FILTER RECONCILIATION ---
        print("Phase 2: Reconciling filters")

        for src_vt in selected_templates:
            for filt_id in src_vt.GetFilters():
                src_filter = src_doc.GetElement(filt_id)
                if not src_filter:
                    continue

                name = src_filter.Name
                src_sig = filter_signature(src_filter)

                if name in dest_filters:
                    dest_filter = dest_filters[name]
                    dest_sig = filter_signature(dest_filter)

                    if src_sig != dest_sig:
                        logger.debug("  Filter differs, updating: {}".format(name))
                        update_filter_in_place(dest_filter, src_filter)
                    else:
                        logger.debug("  Filter identical, reusing: {}".format(name))
                else:
                    logger.debug("  Filter missing, will be copied: {}".format(name))

        # --- COPY TEMPLATES ---
        print("Phase 3: Copying view templates")

        src_ids = List[DB.ElementId]([vt.Id for vt in selected_templates])

        new_ids = DB.ElementTransformUtils.CopyElements(
            src_doc,
            src_ids,
            dest_doc,
            None,
            DB.CopyPasteOptions()
        )

        print("  Copied {} templates".format(len(new_ids)))

        # --- REASSIGN / DELETE / RENAME ---
        print("Phase 4: Overwrite existing templates")

        for new_id in new_ids:
            new_vt = dest_doc.GetElement(new_id)
            if not new_vt:
                continue

            new_name = new_vt.Name

            if new_name.endswith(" 1"):
                base_name = new_name[:-2]

                if base_name in dest_templates:
                    logger.debug("  Overwriting template: {}".format(base_name))

                    # Reassign views
                    for v in template_usage.get(base_name, []):
                        v.ViewTemplateId = new_id
                        logger.debug("    Reassigned view: {}".format(v.Name))

                    # Delete old template
                    dest_doc.Delete(dest_templates[base_name].Id)
                    logger.debug("    Deleted old template")

                    # Rename new template
                    new_vt.Name = base_name
                    logger.debug("    Renamed new template")

                else:
                    logger.debug("  Renaming new template (no conflict)")
                    new_vt.Name = base_name

        print("Completed document: {}".format(dest_doc.Title))

output.print_md("### ✅ View Template Overwrite Complete")
