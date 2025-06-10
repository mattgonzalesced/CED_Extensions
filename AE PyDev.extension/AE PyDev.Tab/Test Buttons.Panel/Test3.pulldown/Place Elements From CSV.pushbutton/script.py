# -*- coding: utf-8 -*-

from pyrevit import DB
from pyrevit import forms
from pyrevit import revit
from pyrevit import script

doc = revit.doc
uidoc = revit.uidoc

# Initialize output window
output = script.get_output()
logger = script.get_logger()

def print_filter_rules(filter_elem):
    """Prints the filter rules of a given filter element."""
    rules = filter_elem.GetElementFilterParameters()
    if not rules:
        logger.info("No filter rules found in: {}".format(filter_elem.Name))
        return

    logger.info("Filter rules for '{}':".format(filter_elem.Name))
    for idx, rule in enumerate(rules):
        try:
            param_id = rule.ParameterId
            param_elem = doc.GetElement(param_id)
            param_name = param_elem.Name if param_elem else str(param_id.IntegerValue)
            logger.info("  Rule {}: Parameter = '{}', Rule Type = '{}'".format(
                idx+1,
                param_name,
                rule.GetType().Name
            ))
        except Exception as e:
            logger.warning("  Error processing rule {}: {}".format(idx+1, e))

def main():
    # Prompt the user to select a filter element
    filter_collector = DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement)
    filters = list(filter_collector)
    if not filters:
        forms.alert("No filters found in the project.", exitscript=True)

    selected_filter = forms.SelectFromList.show(
        [f.Name for f in filters],
        title="Select a View Filter"
    )

    if not selected_filter:
        forms.alert("No filter selected.", exitscript=True)

    # Find the filter element by name
    filter_elem = next((f for f in filters if f.Name == selected_filter), None)
    if filter_elem:
        print_filter_rules(filter_elem)
    else:
        forms.alert("Selected filter not found.", exitscript=True)

if __name__ == "__main__":
    main()