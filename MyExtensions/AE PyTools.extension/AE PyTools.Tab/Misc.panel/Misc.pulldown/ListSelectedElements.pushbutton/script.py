# -*- coding: utf-8 -*-
__title__ = "List Selected Elements"

from pyrevit import script, revit, DB, forms

output = script.get_output()

# Global variable for selected elements
selection = revit.get_selection()


def pick_categories(elements):
    """Prompt the user to select categories if more than one is available."""
    categories = {el.Category.Name for el in elements if el.Category}
    if len(categories) > 1:
        selected_categories = forms.SelectFromList.show(
            sorted(categories),
            title="Select Categories to Include",
            button_name="OK",
            multiselect=True
        )
        if not selected_categories:
            script.exit()
        return [el for el in elements if el.Category and el.Category.Name in selected_categories]
    return elements


def pick_parameters(elements):
    """Gather unique parameters from both instance and type elements, then prompt the user to select parameters."""
    parameters = set()
    for element in elements:
        parameters.update(
            {param.Definition.Name for param in element.Parameters if param.Definition.Name.lower() != "category"})

        # Get type parameters
        element_type = revit.doc.GetElement(element.GetTypeId())
        if element_type:
            parameters.update({param.Definition.Name for param in element_type.Parameters if
                               param.Definition.Name.lower() != "category"})

    sorted_params = sorted(parameters)
    selected_params = forms.SelectFromList.show(
        sorted_params,
        title="Select Parameters to Display",
        button_name="OK",
        multiselect=True
    )




    if not selected_params:
        script.exit()
    return selected_params


def retrieve_parameter_value(element, param_name):
    """Retrieve the value of a parameter, handling different cases for 'n/a' and empty strings."""
    param = element.LookupParameter(param_name)

    if not param:
        # Parameter does not exist on this element
        return "n/a"

    if param.HasValue:
        return format_parameter_value(param)

    # If parameter exists but has no value, return an empty string
    return ""


def format_parameter_value(param):
    """Format the parameter value based on its storage type."""
    try:
        # Use AsValueString if possible
        return param.AsValueString()
    except:
        if param.StorageType == DB.StorageType.String:
            return param.AsString()
        elif param.StorageType == DB.StorageType.Integer:
            return param.AsValueString()
        elif param.StorageType == DB.StorageType.Double:
            return str(DB.UnitUtils.ConvertFromInternalUnits(param.AsDouble(), param.DisplayUnitType))
        elif param.StorageType == DB.StorageType.ElementId:
            # Retrieve and display the name of the linked element
            linked_element = revit.doc.GetElement(param.AsElementId())
            return linked_element.Name if linked_element else str(param.AsElementId().IntegerValue)
        return ""


def collect_element_data(elements, selected_params):
    """Collect data for each element without updating the output until all processing is complete."""
    element_data = []
    for element in elements:
        element_id = output.linkify(DB.ElementId(element.Id.IntegerValue))
        category_name = element.Category.Name if element.Category else "N/A"
        row = [element_id, category_name]

        # Collect parameter values for the row
        row.extend(retrieve_parameter_value(element, param) for param in selected_params)
        element_data.append((category_name, row))

    return element_data


def print_report(element_data, columns):
    """Prints the report title, legend, and table in one go after processing completes."""
    # Print title and legend
    output.print_md("## Element Parameter Report")
    output.print_md("**Legend**")
    output.print_md("- `n/a`: Parameter does not exist for this element.")

    output.insert_divider()

    # Sort data by category and print table
    element_data.sort(key=lambda x: x[0])
    table_rows = [row[1] for row in element_data]
    output.print_table(table_data=table_rows, columns=columns)


# Define the main steps in the script
max_steps = 5  # Number of key steps

with forms.ProgressBar(title="Initializing...", max_value=max_steps, steps=1) as pb:
    # Step 1: Select Categories
    pb.title = "Selecting Categories... (1/5)"
    pb.update_progress(1, max_steps)
    filtered_elements = pick_categories(selection)

    # Step 2: Select Parameters
    pb.title = "Selecting Parameters... (2/5)"
    pb.update_progress(2, max_steps)
    selected_parameters = pick_parameters(filtered_elements)
    columns = ["Element ID", "Category"] + selected_parameters

    # Step 3: Collect Element Data
    pb.title = "Collecting Element Data... (3/5)"
    pb.update_progress(3, max_steps)
    element_data = collect_element_data(filtered_elements, selected_parameters)

    # Step 4: Preparing Report Content
    pb.title = "Preparing Report Content... (4/5)"
    pb.update_progress(4, max_steps)
    # Processing report content preparation could be an intermediate step if needed

    # Step 5: Printing Report
    pb.title = "Printing Report... (Step 5 of 5)"
    print_report(element_data, columns)
    pb.update_progress(5, max_steps)