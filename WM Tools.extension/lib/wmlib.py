# -*- coding: utf-8 -*-
import re

from pyrevit import script, revit, DB, forms
from pyrevit.interop import xl as pyxl
from pyrevit.revit import query

output = script.get_output()

logger = script.get_logger()


class ParentElement:
    """Store details about a parent (reference) element."""

    def __init__(self, element_id, location_point, facing_orientation, doc):
        self.element_id = element_id
        self.location_point = location_point
        self.facing_orientation = facing_orientation
        self.doc = doc

    @property
    def circuit_number(self):
        return self.get_parameter_value("Circuit #")

    @classmethod
    def from_family_instance(cls, element, doc):
        if not isinstance(element, DB.FamilyInstance):
            logger.debug("Input is not a FamilyInstance: {}".format(element.Id))
            return None
        loc = element.Location
        if not isinstance(loc, DB.LocationPoint):
            logger.debug("Skipping element without valid LocationPoint: {}".format(element.Id))
            return None
        orientation = element.FacingOrientation if hasattr(element, "FacingOrientation") else None
        return cls(element.Id, loc.Point, orientation,doc)

    def get_parameter_value(self, name):

        param = self.doc.GetElement(self.element_id).LookupParameter(name)
        if not param or not param.HasValue:
            return None
        st = param.StorageType
        if st == DB.StorageType.String:
            return param.AsString()
        if st == DB.StorageType.Double:
            return param.AsDouble()
        if st == DB.StorageType.Integer:
            return param.AsInteger()
        return None


class ChildGroup:
    def __init__(self, parent, group_type, doc, offset_distance=2.0):
        self.parent = parent
        self.group_type = group_type
        self.offset_distance = offset_distance
        self.orientation = parent.facing_orientation
        self.child_id = None
        self.doc = doc
        self.location = self.offset_placement_point()


    def offset_placement_point(self):
        base_location = self.parent.location_point
        norm_orientation = DB.XYZ.Normalize(self.orientation)
        offset_vector = norm_orientation.Multiply(self.offset_distance)
        offset_location = base_location.Add(offset_vector)
        return offset_location

    def place(self):
        inst = self.doc.Create.PlaceGroup(self.location, self.group_type)
        self.child_id = inst.Id
        logger.info("Placed model-group ID {}".format(self.child_id))
        return inst

    def rotate_to_match_parent(self):
        if not self.child_id or not self.orientation:
            return False
        default = DB.XYZ(0, 1, 0)
        angle = default.AngleTo(self.orientation)
        if default.CrossProduct(self.orientation).Z < 0:
            angle = -angle
        axis = DB.Line.CreateBound(
            self.location,
            self.location + DB.XYZ(0, 0, 1)
        )
        try:
            grp = self.doc.GetElement(self.child_id)
            grp.Location.Rotate(axis, angle)
            logger.info("Rotated model-group ID {} by {}".format(self.child_id, angle))
            return True
        except Exception as e:
            logger.error("Rotation failed: {}".format(e))
            return False

    def copy_parameters(self, mapping):
        grp = self.doc.GetElement(self.child_id)
        for p_name, c_name in mapping.items():
            val = self.parent.get_parameter_value(p_name)
            if val is not None:
                param = grp.LookupParameter(c_name)
                if param and not param.IsReadOnly:
                    try:
                        st = param.StorageType
                        if st == DB.StorageType.String:
                            param.Set(str(val))
                        elif st == DB.StorageType.Double:
                            param.Set(float(val))
                        elif st == DB.StorageType.Integer:
                            param.Set(int(val))
                        elif st == DB.StorageType.ElementId:
                            param.Set(val)
                    except Exception as e:
                        logger.error("Setting {} failed: {}".format(c_name, e))

    def attach_detail_group_by_type(self, detail_group_type):
        if not self.child_id:
            logger.warning("Model group not placed yet.")
            return False
        try:
            group = self.doc.GetElement(self.child_id)
            group.ShowAttachedDetailGroups(self.doc.ActiveView, detail_group_type.Id)
            logger.info(
                "Attached detail group '{}' to group {}".format(query.get_name(detail_group_type), self.child_id))
            return True
        except Exception as e:
            logger.error("Failed to attach detail group '{}': {}".format(query.get_name(detail_group_type), e))
            return False

    @classmethod
    def from_existing_group(cls, group, param_name,doc):
        if not isinstance(group, DB.Group):
            logger.error("Provided element is not a Group: {}".format(group.Id))
            return None
        loc = group.Location
        if not isinstance(loc, DB.LocationPoint):
            logger.warning("Group {} has no LocationPoint".format(group.Id))
            return None

        parent_stub = type('ParentStub', (), {
            'location_point': loc.Point,
            'facing_orientation': DB.XYZ(0, 1, 0),
            'get_parameter_value': lambda self, name: group.LookupParameter(name).AsString() if group.LookupParameter(
                name) else None
        })()

        instance = cls(parent_stub, group.GroupType,doc)
        instance.child_id = group.Id
        return instance

    def ungroup_and_propagate(self, circuit_param, system_param):
        group = self.doc.GetElement(self.child_id)
        if not isinstance(group, DB.Group):
            logger.warning("Element is not a group: {}".format(self.child_id))
            return []

        circuit_number = None
        param = group.LookupParameter(circuit_param)
        if param and param.HasValue:
            circuit_number = param.AsString()

        if not circuit_number:
            logger.warning("Group {} has no circuit number.".format(group.Id))
            return []

        system_number = extract_system_number(circuit_number)

        try:
            self.doc.Regenerate()
        except Exception as e:
            logger.error("Failed to regenerate before ungrouping: {}".format(e))

        ungrouped_ids = []

        def is_attached_to_group(x):
            return hasattr(x, "AttachedParentId") and x.AttachedParentId == group.Id

        detail_groups = DB.FilteredElementCollector(self.doc, self.doc.ActiveView.Id) \
            .OfCategory(DB.BuiltInCategory.OST_IOSAttachedDetailGroups) \
            .WhereElementIsNotElementType().ToElements()

        for dg in filter(is_attached_to_group, detail_groups):
            try:
                ids = dg.UngroupMembers()
                ungrouped_ids.extend(ids)
                logger.info("Ungrouped detail group {}".format(dg.Id))
            except Exception as e:
                logger.error("Failed to ungroup detail group {}: {}".format(dg.Id, e))

        try:
            model_ids = group.UngroupMembers()
            ungrouped_ids.extend(model_ids)
            logger.info("Ungrouped model group {}".format(group.Id))
        except Exception as e:
            logger.error("Failed to ungroup model group {}: {}".format(group.Id, e))

        for eid in ungrouped_ids:
            el = self.doc.GetElement(eid)
            if isinstance(el, DB.FamilyInstance):
                if el.Category and el.Category.Id.Value == int(DB.BuiltInCategory.OST_ElectricalFixtures):
                    p1 = el.LookupParameter(circuit_param)
                    if p1 and not p1.IsReadOnly:
                        try:
                            p1.Set(str(circuit_number))
                            logger.info("Wrote circuit number to fixture ID {}".format(eid))
                        except Exception as e:
                            logger.error("Failed to set circuit number on ID {}: {}".format(eid, e))
                    if system_number:
                        p2 = el.LookupParameter(system_param)
                        if p2 and not p2.IsReadOnly:
                            try:
                                p2.Set(str(system_number))
                                logger.info("Wrote system number to fixture ID {}".format(eid))
                            except Exception as e:
                                logger.error("Failed to set system number on ID {}: {}".format(eid, e))

    @classmethod
    def collect_target_groups(cls, group_type_name):
        provider = DB.ParameterValueProvider(DB.ElementId(DB.BuiltInParameter.SYMBOL_NAME_PARAM))
        evaluator = DB.FilterStringEquals()
        rule = DB.FilterStringRule(provider, evaluator, group_type_name)
        filter_ = DB.ElementParameterFilter(rule)

        return DB.FilteredElementCollector(revit.doc) \
            .OfClass(DB.Group) \
            .WherePasses(filter_) \
            .ToElements()


def collect_reference_tags(doc):
    selected_ids = revit.get_selection().element_ids
    if selected_ids:
        selected_elements = [revit.doc.GetElement(eid) for eid in selected_ids]
        tags = [
            inst for inst in selected_elements
            if isinstance(inst, DB.FamilyInstance)
               and inst.Symbol.Family.Name == "Refrigeration Case Tag - EMS"
               and query.get_name(inst.Symbol) == "EMS Circuit Label"
        ]
        if tags:
            logger.info("Using {} selected EMS tags.".format(len(tags)))
            return tags
        else:
            logger.warning("Selection has no matching EMS tags; falling back to view scan.")

    view_id = doc.ActiveView.Id
    collector = DB.FilteredElementCollector(revit.doc, view_id).OfClass(DB.FamilyInstance)
    tags = [
        inst for inst in collector
        if inst.Symbol.Family.Name == "Refrigeration Case Tag - EMS"
           and query.get_name(inst.Symbol) == "EMS Circuit Label"
    ]
    if not tags:
        logger.error("No matching refrigeration tags found in active view.")
        script.exit()
    return tags


def get_model_group_type(name):
    for gt in DB.FilteredElementCollector(revit.doc).OfClass(DB.GroupType):
        if query.get_name(gt) == name:
            return gt
    logger.error("GroupType not found: {!r}".format(name))
    script.exit()


def get_attached_detail_types(group_type):
    detail_types = []
    for dt_id in group_type.GetAvailableAttachedDetailGroupTypeIds():
        dt = revit.doc.GetElement(dt_id)
        detail_types.append(dt)
    return detail_types


def extract_system_id(circuit_number):
    if not circuit_number:
        return None
    match = re.match(r"^([A-Z]+\d+)", circuit_number)
    if match:
        return match.group(1)
    return None


def extract_system_number(circuit_number):
    if not circuit_number:
        return None
    match = re.match(r'^([^a-z]+)', circuit_number)
    if match:
        return match.group(1)
    return None


class SupressWarnings(DB.IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        ignored_fails = [
            DB.BuiltInFailures.ElectricalFailures.CircuitOverload,
            DB.BuiltInFailures.OverlapFailures.DuplicateInstances
        ]

        try:
            failures = failuresAccessor.GetFailureMessages()

            for fail in failures:  # type: DB.FailureMessageAccessor
                severity = fail.GetSeverity()
                description = fail.GetDescriptionText()
                fail_id = fail.GetFailureDefinitionId()

                if severity == DB.FailureSeverity.Warning and fail_id in ignored_fails:
                    print('✅ Suppressed Warning: {}'.format(description))
                    failuresAccessor.DeleteWarning(fail)

        except Exception as e:
            print('⚠️ Exception in SuppressWarnings: {}'.format(e))

        return DB.FailureProcessingResult.Continue


class ExcelCircuitLoader(object):
    def __init__(self):
        self.path = None
        self.data = {}
        self.ignored_sheets = ["References", "Panel Creation"]
        self.required_headers = [
            "CKT_Panel_CEDT", "CKT_Circuit Number_CEDT", "CKT_Load Name_CEDT",
            "CKT_Rating_CED", "CKT_Frame_CED", "Voltage_CED", "Number of Poles_CED",
            "Apparent Load Ph 1_CED", "Apparent Load Ph 2_CED", "Apparent Load Ph 3_CED","Apparent Load Input_CED"
            "Family", "Type"
        ]

    def pick_excel_file(self):
        self.path = forms.pick_file(multi_file=False)
        if not self.path:
            forms.alert("No Excel file selected.")
            return

        self.data = pyxl.load(self.path, headers=False)

    def get_valid_sheet_names(self):
        return [s for s in self.data.keys() if s not in self.ignored_sheets]

    def build_dict_rows(self):
        for sheetname, sheet in self.data.items():
            logger.debug("🔍 Processing sheet: **{}**".format(sheetname))
            rows = sheet.get("rows", [])
            if len(rows) < 2:
                output.print_md("⚠️ Sheet '{}' has less than 2 rows.".format(sheetname))
                continue

            header_row = [str(h).strip() if h is not None else "" for h in rows[0]]
            logger.debug("📋 Headers from '{}':\n{}".format(sheetname, ", ".join(header_row)))

            dict_rows = []
            for i, values in enumerate(rows[1:], start=2):  # Start from second row (row 2)
                row_dict = {}

                for j in range(len(header_row)):
                    col = header_row[j]
                    val = values[j] if j < len(values) else None
                    row_dict[col] = val


                dict_rows.append(row_dict)

            self.data[sheetname] = {
                "columns": header_row,
                "rows": dict_rows
            }

    def pick_sheet_names(self, sheetnames):
        return forms.SelectFromList.show(sheetnames,
                                         title="Select Circuit Sheets",
                                         multiselect=True)

    def validate_row_headers(self, row):
        missing = [h for h in self.required_headers if h not in row]
        if missing:
            logger.warning("Row missing expected headers: {}".format(", ".join(missing)))
            return False
        return True

    def get_ordered_rows(self, sheetnames):
        ordered_rows = []

        for sheet in sheetnames:
            for row in self.data.get(sheet, {}).get("rows", []):
                if not isinstance(row, dict):
                    continue

                clean_row = {}
                for k, v in row.items():
                    if isinstance(v, str):
                        clean_row[k] = v.strip()
                    elif k in ["Voltage_CED", "Number of Poles_CED"] and isinstance(v, (int, float)):
                        clean_row[k] = int(v)
                    else:
                        clean_row[k] = v

                # ✅ Safely handle CKT_Panel_CEDT and CKT_Circuit Number_CEDT
                panel_val = clean_row.get("CKT_Panel_CEDT")
                circuit_val = clean_row.get("CKT_Circuit Number_CEDT")
                panel = str(panel_val).strip() if isinstance(panel_val, basestring) else str(panel_val)
                circuit = str(circuit_val).strip() if isinstance(circuit_val, basestring) else str(circuit_val)

                if not (panel and circuit):
                    continue

                clean_row = self.apply_family_type_fallback(clean_row)

                if clean_row.get("Family") and clean_row.get("Type"):
                    ordered_rows.append(clean_row)

        output.print_md("✅ Final ordered row count: **{}**".format(len(ordered_rows)))
        return ordered_rows

    def clean_row_data(self, row):
        clean_row = {}
        for k, v in row.items():
            if isinstance(v, str):
                clean_row[k] = v.strip()
            elif k in ["Voltage_CED", "Number of Poles_CED"] and isinstance(v, (int, float)):
                clean_row[k] = int(v)
            else:
                clean_row[k] = v
        return clean_row

    def apply_family_type_fallback(self, row):
        type_map = {
            (120, 1): "120V/1P",
            (208, 2): "208V/2P",
            (208, 3): "208V/3P",
            (277, 1): "277V/1P",
            (480, 2): "480V/2P",
            (480, 3): "480V/3P",
        }
        fallback_family = "EF-F_Existing Ckt Placeholder-Unbalanced_CED"

        family = row.get("Family", "").strip()
        typ = row.get("Type", "").strip()

        if family and typ:
            return row  # already valid

        try:
            voltage = int(float(row.get("Voltage_CED", 0)))
            poles = int(float(row.get("Number of Poles_CED", 0)))
            fallback_type = type_map.get((voltage, poles))

            if fallback_type:
                row["Family"] = fallback_family
                row["Type"] = fallback_type

            else:
                logger.warning("⚠️ No fallback match for Voltage={} Poles={}".format(voltage, poles))
        except Exception as e:
            logger.warning("❌ Invalid voltage or poles in row: {}".format(row))

        return row


class EquipmentSurface(object):
    def __init__(self, element_id):
        self.element_id = element_id
        self.element = revit.doc.GetElement(DB.ElementId(element_id))
        self.name = self._get_panel_name()
        self.location = self._get_location_point()
        self.facing = self._get_facing_orientation()
        self.face = None
        self.normal = None
        self._resolve_geometry()



    def _get_panel_name(self):
        param = self.element.LookupParameter("Panel Name_CEDT")
        return param.AsString() if param and param.HasValue else None

    def _get_location_point(self):
        loc = self.element.Location
        return loc.Point if isinstance(loc, DB.LocationPoint) else None

    def _get_facing_orientation(self):
        return self.element.FacingOrientation if hasattr(self.element, 'FacingOrientation') else None

    def _resolve_geometry(self):
        opt = DB.Options()
        opt.ComputeReferences = True  # ✅ Required!
        opt.View = revit.doc.ActiveView
        geom = self.element.get_Geometry(opt)
        for geo_obj in geom:
            if isinstance(geo_obj, DB.GeometryInstance):
                inst_geom = geo_obj.GetSymbolGeometry()
                for g in inst_geom:
                    if isinstance(g, DB.Solid) and g.Faces.Size > 0:
                        style_id = g.GraphicsStyleId
                        if not style_id or style_id.Value < 0:
                            continue
                        style = revit.doc.GetElement(style_id)
                        if not style:
                            continue
                        if style.Name not in ["Panelboards_CED", "Switchboards_CED"]:
                            continue
                        for f in g.Faces:
                            if isinstance(f, DB.PlanarFace) and f.FaceNormal.IsAlmostEqualTo(DB.XYZ.BasisZ):
                                self.face = f
                                self.normal = f.FaceNormal
                                return
