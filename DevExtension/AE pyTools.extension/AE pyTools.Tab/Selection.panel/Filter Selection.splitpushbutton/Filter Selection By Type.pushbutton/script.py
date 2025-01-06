# -*- coding: utf-8 -*-
from pyrevit import revit, DB
from pyrevit import forms

doc = revit.doc
uidoc = revit.uidoc


# Function to group elements by category and family types
def group_elements_by_types(selection):
    types_by_category = {}

    for element in selection:
        category = element.Category.Name if element.Category else "<No Category>"
        family_symbol = element.GetTypeId()
        family_name = "<No Family>"
        type_name = "<No Type>"

        if family_symbol and family_symbol != DB.ElementId.InvalidElementId:
            family_symbol_element = revit.doc.GetElement(family_symbol)
            if family_symbol_element:
                family_name = family_symbol_element.FamilyName
                type_name = DB.Element.Name.__get__(family_symbol_element)

        key = (category, family_name, type_name)
        if category not in types_by_category:
            types_by_category[category] = {}
        if key not in types_by_category[category]:
            types_by_category[category][key] = []
        types_by_category[category][key].append(element)

    return types_by_category


# Function to prompt user for filtering selection
def prompt_user_for_selection(types_by_category):
    # Prepare options for user prompt
    grouped_options = {' All': []}
    for category, groups in types_by_category.items():
        if category not in grouped_options:
            grouped_options[category] = []
        for (cat, family_name, type_name), elements in groups.items():
            count = len(elements)
            option_text = "{} | {} | {} ({})".format(cat, family_name, type_name, count)
            grouped_options[' All'].append(option_text)
            grouped_options[category].append(option_text)

    # Sort options for consistent display
    for key in grouped_options:
        grouped_options[key].sort()

    # Prompt user to select the desired types with a group selector
    selected_options = forms.SelectFromList.show(
        grouped_options,
        title="Filter Selection By Type",
        group_selector_title="Category:",
        multiselect=True


    )

    if not selected_options:
        forms.alert("No types selected. Operation cancelled.", exitscript=True)

    # Collect elements matching selected types
    filtered_elements = set()
    for category, groups in types_by_category.items():
        for key, elements in groups.items():
            if "{} | {} | {} ({})".format(key[0], key[1], key[2], len(elements)) in selected_options:
                filtered_elements.update(elements)

    return filtered_elements


# Main function
def main():
    selection = revit.get_selection()
    selection_ids = selection.element_ids

    if not selection_ids:
        forms.alert("Please select some elements before running this tool.", exitscript=True)

    selection_elements = [doc.GetElement(id) for id in selection_ids]

    # Group elements by types
    types_by_category = group_elements_by_types(selection_elements)

    # Prompt user and get new filtered set
    filtered_elements = prompt_user_for_selection(types_by_category)

    # Update selection in Revit
    selection.set_to(filtered_elements)


if __name__ == "__main__":
    main()
