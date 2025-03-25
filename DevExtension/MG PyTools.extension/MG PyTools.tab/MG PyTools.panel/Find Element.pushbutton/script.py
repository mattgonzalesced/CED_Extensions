from pyrevit import revit, DB
from pyrevit import script
from pyrevit.forms import ask_for_string, alert  # added alert import

# Initialize output manager
output = script.get_output()
output.close_others()

# Prompt user for Element ID
try:
    input_id_str = ask_for_string("")
    if input_id_str is None:
        raise ValueError("No input provided. Script canceled.")
    input_id = int(input_id_str)
    if input_id <= 0:
        raise ValueError("Element ID must be a positive integer.")
except ValueError as e:
    script.get_logger().error("Invalid Element ID: {0}".format(str(e)))
    raise

# Get the element by ID
element = revit.doc.GetElement(DB.ElementId(input_id))

if element:
    # Get the parent view or owner view (if applicable)
    owner_view_id = None
    parent_view = None

    try:
        if hasattr(element, "OwnerViewId") and element.OwnerViewId != DB.ElementId.InvalidElementId:
            owner_view_id = element.OwnerViewId
            parent_view = revit.doc.GetElement(owner_view_id)
    except Exception as e:
        script.get_logger().warning("Could not retrieve owner view: {0}".format(str(e)))

    # Get element name and category
    try:
        element_name = element.Name if hasattr(element, "Name") else "(No Name)"
    except Exception as e:
        element_name = "(Error retrieving name)"
        script.get_logger().warning("Error retrieving name for element ID {0}: {1}".format(input_id, str(e)))

    category_name = element.Category.Name if element.Category else "(No Category)"

    # Create a clickable link to the element
    try:
        clickable_element_id = output.linkify([element.Id]) if element.Id else "(Invalid ID)"
    except Exception as e:
        clickable_element_id = "(Error creating link)"
        script.get_logger().warning("Error creating link for element ID {0}: {1}".format(input_id, str(e)))

    # Automatically navigate to the element by passing the element itself.
    revit.uidoc.ShowElements(element)
else:
    # If the element is not found, show a popup alert.
    alert("Element with ID {0} not found in the model.".format(input_id))
