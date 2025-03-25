# -*- coding: utf-8 -*-
import clr
import Autodesk.Revit.UI as UI
import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory, IndependentTag, Reference, XYZ
import sys

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
view = doc.ActiveView

def get_tag_symbol_by_family_and_type(tag_symbols, desired_family, desired_type):
    """Search for a matching tag symbol based on family and type names."""
    for tag_symbol in tag_symbols:
        try:
            # Retrieve the Family Name correctly
            family_name = tag_symbol.Family.Name if tag_symbol.Family else None

            # Retrieve the Type Name via Built-in Parameter
            type_param = tag_symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME)
            type_name = type_param.AsString() if type_param else None

            if family_name and type_name:
                if family_name == desired_family and type_name == desired_type:
                    return tag_symbol
        except Exception as e:
            print("DEBUG: Error checking symbol {}: {}".format(tag_symbol.Id, e))
    
    return None

def tag_family_instances(doc, view):
    print("DEBUG: Starting tag_family_instances")

    if view.ViewType != DB.ViewType.DraftingView:
        print("DEBUG: Current view is not a Drafting View. Exiting.")
        return

    # Collect all detail component instances in the active view
    family_instances = list(FilteredElementCollector(doc, view.Id)
                            .OfCategory(BuiltInCategory.OST_DetailComponents)
                            .WhereElementIsNotElementType())
    print("DEBUG: Found {} family instances.".format(len(family_instances)))

    if not family_instances:
        print("DEBUG: No family instances found. Exiting.")
        return

    # Get existing tags (via TAGGED_ELEM_ID) to avoid duplicates
    existing_tags = {tag.TaggedElementId.IntegerValue for tag in 
                     FilteredElementCollector(doc, view.Id)
                     .OfClass(IndependentTag)
                     .WhereElementIsNotElementType()}
    
    print("DEBUG: Found {} existing tags.".format(len(existing_tags)))

    # Get all available tag symbols
    tag_symbols = list(FilteredElementCollector(doc)
                       .OfClass(DB.FamilySymbol)
                       .OfCategory(BuiltInCategory.OST_DetailComponentTags)
                       .WhereElementIsElementType())

    print("DEBUG: Found {} tag symbols.".format(len(tag_symbols)))

    # Define tag mapping rules
    tag_mapping = {
        "DME-EQU-Switchboard-Top_CED": ("MG_DI-Tag_Panelboard-L_CED", "Panelboard - All On"),
        "DME-EQU-Panel-Top_CED": ("MG_DI-Tag_Panelboard-L_CED", "Panelboard - All On")
    }

    # Default mapping if an instance family isn't explicitly mapped
    default_tag = ("MG_DI-Tag_Circuit Breaker_CED", "Panel Name, Voltage, Mains")

    with Transaction(doc, "Tag Family Instances") as t:
        t.Start()
        for instance in family_instances:
            try:
                instance_id = instance.Id.IntegerValue
                instance_family = instance.Symbol.Family.Name
                print("DEBUG: Processing instance {} (Family: {}).".format(instance_id, instance_family))

                if instance_id in existing_tags:
                    print("DEBUG: Instance {} already has a tag. Skipping.".format(instance_id))
                    continue

                reference = Reference(instance)
                location = instance.Location.Point if hasattr(instance, "Location") and isinstance(instance.Location, DB.LocationPoint) else XYZ(0, 0, 0)

                # Determine the desired tag
                tag_family, tag_type = tag_mapping.get(instance_family, default_tag)
                print("DEBUG: For family '{}', using tag Family '{}' Type '{}'.".format(instance_family, tag_family, tag_type))

                # Find the matching tag symbol
                tag_symbol = get_tag_symbol_by_family_and_type(tag_symbols, tag_family, tag_type)

                if tag_symbol:
                    if not tag_symbol.IsActive:
                        tag_symbol.Activate()
                        doc.Regenerate()
                    IndependentTag.Create(doc, tag_symbol.Id, view.Id, reference, False, DB.TagOrientation.Horizontal, location)
                    print("DEBUG: Tag created for instance {}.".format(instance_id))
                else:
                    print("DEBUG: No matching tag found for instance {} with Family '{}' and Type '{}'.".format(instance_id, tag_family, tag_type))

            except Exception as e:
                print("DEBUG: Exception processing instance {}: {}".format(instance.Id.IntegerValue, e))
                continue

        t.Commit()
        print("DEBUG: Transaction committed.")

tag_family_instances(doc, view)
