import clr
import Autodesk.Revit.UI as UI
import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory, IndependentTag, Reference, XYZ
import sys

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
view = doc.ActiveView

def tag_family_instances(doc, view):
    """Automatically tags family instances in a drafting view."""
    
    # Ensure the script runs only in a Drafting View
    if view.ViewType != DB.ViewType.DraftingView:
        return
    
    # Collect all detail components in the active view
    family_instances = list(FilteredElementCollector(doc, view.Id)
                            .OfCategory(BuiltInCategory.OST_DetailComponents)
                            .WhereElementIsNotElementType())
    if not family_instances:
        return

    # Get existing tags to avoid duplication
    existing_tags = {tag.TaggedElementId.IntegerValue for tag in FilteredElementCollector(doc, view.Id)
                     .OfClass(IndependentTag).WhereElementIsNotElementType()}

    # Get all available detail item tags
    detail_item_tags = {tag.Family.Name: tag for tag in FilteredElementCollector(doc)
                        .OfClass(DB.FamilySymbol)
                        .OfCategory(BuiltInCategory.OST_DetailComponentTags)}
    if not detail_item_tags:
        return

    # Define tag mapping rules
    tag_mapping = {
        "DME-EQU-Switchboard-Top_CED": ("MG_DI-Tag_Panelboard-C_CED", "Panelboard - All On"),
        "DME-EQU-Transformer-Box-Ground-Top_CED": ("MG_DI-Tag_Panelboard-C_CED", "Panelboard - All On"),
        "DME-EQU-Panel-Top_CED": ("MG_DI-Tag_Panelboard-C_CED", "Panelboard - All On")
    }
    default_tag = ("MG_DI-Tag_Circuit Breaker_CED", "Panel Name, Voltage, Mains")

    # Start a transaction to place tags
    with Transaction(doc, "Tag Family Instances") as t:
        t.Start()
        for instance in family_instances:
            if instance.Id.IntegerValue in existing_tags:
                continue
            try:
                reference = Reference(instance)
                location = instance.Location.Point if isinstance(instance.Location, DB.LocationPoint) else XYZ(0, 0, 0)
                
                # Select appropriate tag
                family_name = instance.Symbol.Family.Name
                tag_family, tag_type = tag_mapping.get(family_name, default_tag)
                
                if tag_family in detail_item_tags:
                    tag_type_id = detail_item_tags[tag_family].Id
                    IndependentTag.Create(doc, tag_type_id, view.Id, reference, False, DB.TagOrientation.Horizontal, location)
            except:
                continue
        t.Commit()

# Call the function
tag_family_instances(doc, view)
