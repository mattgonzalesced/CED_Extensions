# -*- coding: UTF-8 -*-
from Autodesk.Revit.DB import (
    BuiltInParameter,
    Level,
    MEPCurve,
    FamilyInstance,
    FilteredElementCollector,
    BuiltInCategory,
    Category,
    StorageType,
    FamilyPlacementType,

)
from pyrevit import script, revit
from pyrevit.forms import WPFWindow

from pyrevitmep.event import CustomizableEvent

# Globals
logger = script.get_logger()
doc = revit.doc
uidoc = revit.uidoc

# --- ELEVATION MODES ---
ELEVATION_MATCH_LEVEL_ONLY = 0 # Level 2, maintain current elevation
ELEVATION_FROM_LEVEL = 1 # Level 2, +5.0 ft above level
ELEVATION_BBOX_TOP = 2 # Level 2, +15.0 ft above level
ELEVATION_BBOX_MIDDLE = 3 # Level 2, +12.5 ft above level
ELEVATION_BBOX_BOTTOM = 4 # Level 2,  +5.0 ft above level

# --- OFFSET LOGIC WRAPPER ---
class OffsetManager(object):
    def __init__(self, ref_level, elevation_mode, reference_element):
        self.ref_level = ref_level
        self.elevation_mode = elevation_mode
        self.reference_element = reference_element
        self.offsettable_ids = self._get_category_ids(self._get_offsettable_categories())
        self.target_elevation = None if elevation_mode == ELEVATION_MATCH_LEVEL_ONLY else self._get_target_elevation()

    def _get_offsettable_categories(self):
        return {
            BuiltInCategory.OST_DuctTerminal,
            BuiltInCategory.OST_AudioVisualDevices,
            BuiltInCategory.OST_CommunicationDevices,
            BuiltInCategory.OST_DataDevices,
            BuiltInCategory.OST_Columns,
            BuiltInCategory.OST_ElectricalEquipment,
            BuiltInCategory.OST_ElectricalFixtures,
            BuiltInCategory.OST_FireAlarmDevices,
            BuiltInCategory.OST_FireProtection,
            BuiltInCategory.OST_GenericModel,
            BuiltInCategory.OST_LightingDevices,
            BuiltInCategory.OST_LightingFixtures,
            BuiltInCategory.OST_MechanicalControlDevices,
            BuiltInCategory.OST_MechanicalEquipment,
            BuiltInCategory.OST_MedicalEquipment,
            BuiltInCategory.OST_NurseCallDevices,
            BuiltInCategory.OST_PlumbingEquipment,
            BuiltInCategory.OST_PlumbingFixtures,
            BuiltInCategory.OST_SecurityDevices,
            BuiltInCategory.OST_SpecialityEquipment,
            BuiltInCategory.OST_Sprinklers,
            BuiltInCategory.OST_TelephoneDevices,
            BuiltInCategory.OST_IOSModelGroups
        }

    def _get_category_ids(self, categories):
        return tuple(Category.GetCategory(doc, cat).Id for cat in categories)

    def _get_target_elevation(self):
        if not self.reference_element:
            return None
        bbox = self.reference_element.get_BoundingBox(None)
        if not bbox:
            return self.ref_level.Elevation

        min_z = bbox.Min.Z
        max_z = bbox.Max.Z
        mid_z = (min_z + max_z) / 2.0

        if self.elevation_mode == ELEVATION_FROM_LEVEL:
            # Use actual elevation offset from level
            elevation_param = self.reference_element.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM)
            if elevation_param:
                return elevation_param.AsDouble() + self.ref_level.Elevation
            else:
                logger.warning("Reference element missing INSTANCE_ELEVATION_PARAM")
                return self.ref_level.Elevation
        elif self.elevation_mode == ELEVATION_BBOX_TOP:
            return max_z
        elif self.elevation_mode == ELEVATION_BBOX_MIDDLE:
            return mid_z
        elif self.elevation_mode == ELEVATION_BBOX_BOTTOM:
            return min_z
        else:
            return None

    def needs_offset(self, element):
        if not isinstance(element, FamilyInstance):
            return False
        if element.Category.Id not in self.offsettable_ids:
            return False
        symbol = element.Symbol
        if symbol and symbol.Family.FamilyPlacementType == FamilyPlacementType.WorkPlaneBased:
            return False
        return True

    def apply_to(self, element):
        if isinstance(element, MEPCurve):
            element.ReferenceLevel = self.ref_level
            return

        if isinstance(element, FamilyInstance):
            el_level = doc.GetElement(element.LevelId)
            placement_type = element.Symbol.Family.FamilyPlacementType

            # Get appropriate level parameter (if one exists)
            if placement_type == FamilyPlacementType.WorkPlaneBased:
                level_param = element.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM)
            else:
                level_param = element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)

            # Fallback to element.LevelId if needed
            if (not level_param) or level_param.IsReadOnly:
                try:
                    element.LevelId = self.ref_level.Id
                except:
                    logger.warning("Cannot set LevelId on element ID {}".format(element.Id))
            else:
                level_param.Set(self.ref_level.Id)

            # Offset only if needed
            if self.needs_offset(element):
                offset_param = element.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
                if offset_param and not offset_param.IsReadOnly:
                    if self.target_elevation is not None:
                        offset = self.target_elevation - self.ref_level.Elevation
                    else:
                        offset = offset_param.AsDouble() + el_level.Elevation - self.ref_level.Elevation
                    offset_param.Set(offset)
            return

        # if isinstance(element, Space):
            # point = element.Location.Point
            # newspace = doc.Create.NewSpace(self.ref_level, UV(point.X, point.Y))
            # for param in element.Parameters:
            #     if not param.IsReadOnly:
            #         val = get_param_value(param)
            #         if val:
            #             newspace.LookupParameter(param.Definition.Name).Set(val)
            # newspace.get_Parameter(BuiltInParameter.ROOM_UPPER_LEVEL).Set(self.ref_level.Id)
            # if newspace.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET).AsDouble() <= 0:
            #     newspace.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET).Set(1 / 0.3048)
            # doc.Delete(element.Id)


# --- HELPER ---
def get_param_value(param):
    if param.StorageType == StorageType.Double:
        return param.AsDouble()
    elif param.StorageType == StorageType.ElementId:
        return param.AsElementId()
    elif param.StorageType == StorageType.Integer:
        return param.AsInteger()
    elif param.StorageType == StorageType.String:
        return param.AsString()


# --- COMMAND ---
def apply_level_change(ref_level, elevation_mode=ELEVATION_MATCH_LEVEL_ONLY, reference_element=None):
    selection_ids = uidoc.Selection.GetElementIds()
    offset_mgr = OffsetManager(ref_level, elevation_mode, reference_element)

    with revit.Transaction("Change Level", doc):
        for el_id in selection_ids:
            element = doc.GetElement(el_id)
            offset_mgr.apply_to(element)


# --- UI ENTRY POINT ---
customizable_event = CustomizableEvent()

class ReferenceLevelSelection(WPFWindow):
    def __init__(self, xaml_file_name):
        self.ref_object = None  # <-- Move this to the top
        WPFWindow.__init__(self, xaml_file_name)
        self.levels = FilteredElementCollector(doc).OfClass(Level)
        self.combobox_levels.DataContext = self.levels


    def from_list_click(self, sender, e):
        level = self.combobox_levels.SelectedItem
        customizable_event.raise_event(apply_level_change, level, ELEVATION_MATCH_LEVEL_ONLY, None)

    def from_object_click(self, sender, e):
        self.ref_object = revit.pick_element("Pick reference object")
        ref_el = self.ref_object

        # Get Level safely
        level_id = getattr(ref_el, "LevelId", None)
        level = doc.GetElement(level_id) if level_id else None

        # Get name safely
        try:
            name = ref_el.Name
        except:
            name_param = ref_el.get_Parameter(BuiltInParameter.ROOM_NAME)
            name = name_param.AsString() if name_param else "<Unnamed>"

        # Update UI
        category = ref_el.Category.Name if ref_el.Category else "Unknown"
        info_text = "Reference Element:\nCategory: {}\nID: {}\nName: {}".format(category, ref_el.Id, name)
        self.textblock_reference_info.Text = info_text

        self.update_elevation_summary(level)

    def update_elevation_summary(self, ref_level):
        if not self.ref_object:
            return

        mode = self.combo_elevation_mode.SelectedIndex
        offset_mgr = OffsetManager(ref_level, mode, self.ref_object)

        if offset_mgr.target_elevation is None:
            label = "Level: {}, maintain current elevation".format(ref_level.Name)
        else:
            delta_ft = (offset_mgr.target_elevation - ref_level.Elevation)
            label = "Level: {}, +{:.1f} ft above level".format(ref_level.Name, delta_ft)

        self.textblock_mod_summary.Text = label

    def from_object_apply_click(self, sender, e):
        level = doc.GetElement(self.ref_object.LevelId)
        mode = self.combo_elevation_mode.SelectedIndex
        customizable_event.raise_event(apply_level_change, level, mode, self.ref_object)

    def combo_elevation_mode_SelectionChanged(self, sender, e):
        if self.ref_object:
            level = doc.GetElement(self.ref_object.LevelId)
            self.update_elevation_summary(level)


if __forceddebugmode__:
    el = revit.pick_element("Pick reference object")
    lvl = doc.GetElement(el.LevelId)
    apply_level_change(lvl, ELEVATION_MATCH_LEVEL_ONLY, el)
else:
    ReferenceLevelSelection("ReferenceLevelSelection.xaml").Show()
