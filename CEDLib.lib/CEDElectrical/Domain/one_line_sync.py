# -*- coding: utf-8 -*-
"""Sync one-line detail items with model elements."""

from pyrevit import DB, script

logger = script.get_logger()

DETAIL_PARAM_CKT_PANEL = "CKT_Panel_CEDT"
DETAIL_PARAM_CKT_NUMBER = "CKT_Circuit Number_CEDT"
DETAIL_PARAM_PANEL_NAME = "Panel Name_CEDT"
DETAIL_PARAM_SC_PANEL_ID = "SC_Panel ElementId"
DETAIL_PARAM_SC_CIRCUIT_ID = "SC_Circuit ElementId"
SLD_COMPONENT_ID_PARAM = "SLD_Component ID_CED"

DEVICE_CATEGORY_IDS = [
    DB.ElementId(DB.BuiltInCategory.OST_ElectricalFixtures),
    DB.ElementId(DB.BuiltInCategory.OST_LightingFixtures),
    DB.ElementId(DB.BuiltInCategory.OST_DataDevices)
]

CIRCUIT_VALUE_MAP = {
    "x VD Schedule": "x VD Schedule",
    "Circuit Tree Sort_CED": "Circuit Tree Sort_CED",
    "CKT_Circuit Type_CEDT": "CKT_Circuit Type_CEDT",
    "CKT_Panel_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
    "CKT_Circuit Number_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER,
    "CKT_Load Name_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME,
    "CKT_Rating_CED": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM,
    "CKT_Frame_CED": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_FRAME_PARAM,
    "CKT_Schedule Notes_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM,
    "CKT_Length_CED": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_LENGTH_PARAM,
    "Number of Poles_CED": DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES,
    "Voltage_CED": DB.BuiltInParameter.RBS_ELEC_VOLTAGE,
    "Wire Material_CEDT": "Wire Material_CEDT",
    "Wire Insulation_CEDT": "Wire Insulation_CEDT",
    "Wire Temparature Rating_CEDT": "Wire Temparature Rating_CEDT",
    "Wire Size_CEDT": "Wire Size_CEDT",
    "Conduit and Wire Size_CEDT": "Conduit and Wire Size_CEDT",
    "Conduit Type_CEDT": "Conduit Type_CEDT",
    "Conduit Size_CEDT": "Conduit Size_CEDT",
    "Conduit Fill Percentage_CED": "Conduit Fill Percentage_CED",
    "Voltage Drop Percentage_CED": "Voltage Drop Percentage_CED",
    "Circuit Load Current_CED": "Circuit Load Current_CED"
}

PANEL_VALUE_MAP = {
    "Panel Name_CEDT": DB.BuiltInParameter.RBS_ELEC_PANEL_NAME,
    "Mains Rating_CED": "Mains Rating_CED",
    "Mains Type_CEDT": "Mains Type_CEDT",
    "Phase_CED": "Phase_CED",
    "Main Breaker Rating_CED": DB.BuiltInParameter.RBS_ELEC_PANEL_MCB_RATING_PARAM,
    "Short Circuit Rating_CEDT": DB.BuiltInParameter.RBS_ELEC_SHORT_CIRCUIT_RATING,
    "Mounting_CEDT": DB.BuiltInParameter.RBS_ELEC_MOUNTING,
    "Panel Modifications_CEDT": DB.BuiltInParameter.RBS_ELEC_MODIFICATIONS,
    "Distribution System_CEDR": DB.BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM,
    "Secondary Distribution System_CEDR": DB.BuiltInParameter.RBS_FAMILY_CONTENT_SECONDARY_DISTRIBSYS,
    "Total Connected Load_CEDR": DB.BuiltInParameter.RBS_ELEC_PANEL_TOTALLOAD_PARAM,
    "Total Demand Load_CEDR": DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_DEMAND_CURRENT_PARAM,
    "Total Connected Current_CEDR": DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_CONNECTED_CURRENT_PARAM,
    "Total Demand Current_CEDR": DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_DEMAND_CURRENT_PARAM,
    "Max Number of Single Pole Breakers_CED": DB.BuiltInParameter.RBS_ELEC_MAX_POLE_BREAKERS,
    "Max Number of Circuits_CED": DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_CIRCUITS,
    "Transformer Rating_CEDT": "Transformer Rating_CEDT",
    "Transformer Rating_CED": "Transformer Rating_CEDT",
    "Transformer Primary Description_CEDT": "Transformer Primary Description_CEDT",
    "Transformer Secondary Description_CEDT": "Transformer Secondary Description_CEDT",
    "Transformer %Z_CED": "Transformer %Z_CED",
    "Panel Feed_CEDT": DB.BuiltInParameter.RBS_ELEC_PANEL_FEED_PARAM,
}

DEVICE_VALUE_MAP = {
    "CKT_Panel_CEDT": "Panel",
    "CKT_Circuit Number_CEDT": "Circuit Number",
    "CKT_Load Name_CEDT": "Load Name",
    "Voltage_CED": "Voltage",
    "Circuit Load Current_CED": "Load"
}


class SyncAssociation(object):
    def __init__(self, model_elem, detail_elem, kind, key=None):
        self.model_elem = model_elem
        self.detail_elem = detail_elem
        self.kind = kind
        self.key = key
        self.status = None


class OneLineSyncService(object):
    def __init__(self, doc):
        self.doc = doc
        self.option_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)

    # ------------------------------------------------------------------
    # Parameter helpers
    # ------------------------------------------------------------------
    def get_model_param_value(self, elem, param_key, allow_type_fallback=True):
        param = None

        if isinstance(param_key, DB.BuiltInParameter):
            param = elem.get_Parameter(param_key)
        elif isinstance(param_key, str):
            param = elem.LookupParameter(param_key)
        else:
            return None

        if not param and allow_type_fallback:
            try:
                type_elem = elem.Document.GetElement(elem.GetTypeId())
                if type_elem:
                    if isinstance(param_key, DB.BuiltInParameter):
                        param = type_elem.get_Parameter(param_key)
                    elif isinstance(param_key, str):
                        param = type_elem.LookupParameter(param_key)
            except Exception as exc:
                logger.debug("get_model_param_value: Error accessing type element {}: {}".format(elem.Id, exc))

        if not param:
            return None

        st = param.StorageType
        if st == DB.StorageType.String:
            return param.AsString()
        elif st == DB.StorageType.Integer:
            return param.AsInteger()
        elif st == DB.StorageType.Double:
            return param.AsDouble()
        elif st == DB.StorageType.ElementId:
            return param.AsValueString()

        return None

    def get_detail_param_value(self, elem, param_name):
        param = elem.LookupParameter(param_name)
        if not param:
            return None

        st = param.StorageType
        if st == DB.StorageType.String:
            return param.AsString()
        elif st == DB.StorageType.Integer:
            return param.AsInteger()
        elif st == DB.StorageType.Double:
            return param.AsDouble()
        elif st == DB.StorageType.ElementId:
            return param.AsValueString()
        return None

    def set_param_value(self, elem, param_name, new_value):
        param = elem.LookupParameter(param_name)
        if not param or param.IsReadOnly:
            return False

        try:
            if new_value is None:
                if param.StorageType == DB.StorageType.String:
                    param.Set("")
                elif param.StorageType == DB.StorageType.Integer:
                    param.Set(0)
                elif param.StorageType == DB.StorageType.Double:
                    param.Set(0.0)
                return True

            if param.StorageType == DB.StorageType.String:
                param.Set(str(new_value))
            elif param.StorageType == DB.StorageType.Integer:
                param.Set(int(new_value))
            elif param.StorageType == DB.StorageType.Double:
                param.Set(float(new_value))
            elif param.StorageType == DB.StorageType.ElementId and isinstance(new_value, DB.ElementId):
                param.Set(new_value)
            else:
                param.Set(str(new_value))
            return True
        except Exception:
            return False

    def _normalize(self, value):
        if value is None:
            return ""
        return str(value).strip()

    # ------------------------------------------------------------------
    # Collect elements
    # ------------------------------------------------------------------
    def collect_circuits(self):
        return DB.FilteredElementCollector(self.doc) \
            .OfClass(DB.Electrical.ElectricalSystem) \
            .WherePasses(self.option_filter) \
            .ToElements()

    def collect_panels(self):
        return DB.FilteredElementCollector(self.doc) \
            .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment) \
            .WhereElementIsNotElementType() \
            .WherePasses(self.option_filter) \
            .ToElements()

    def collect_devices(self):
        elems = []
        for cat_id in DEVICE_CATEGORY_IDS:
            elems.extend(DB.FilteredElementCollector(self.doc)
                         .OfCategoryId(cat_id)
                         .WhereElementIsNotElementType()
                         .WherePasses(self.option_filter)
                         .ToElements())
        return elems

    def collect_detail_items(self):
        return DB.FilteredElementCollector(self.doc) \
            .OfCategory(DB.BuiltInCategory.OST_DetailComponents) \
            .WhereElementIsNotElementType() \
            .WherePasses(self.option_filter) \
            .ToElements()

    # ------------------------------------------------------------------
    # Detail item associations
    # ------------------------------------------------------------------
    def get_circuit_key(self, circuit):
        panel_val = self.get_model_param_value(circuit, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM)
        cnum_val = self.get_model_param_value(circuit, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER)
        if panel_val and cnum_val:
            return (str(panel_val), str(cnum_val))
        return None

    def get_panel_key(self, panel):
        pname = self.get_model_param_value(panel, DB.BuiltInParameter.RBS_ELEC_PANEL_NAME)
        if pname:
            return str(pname)
        return None

    def get_device_circuit_key(self, device):
        panel_val = None
        cnum_val = None

        try:
            mep = device.MEPModel
            if mep:
                assigned = mep.GetAssignedElectricalSystems()
                if assigned and len(assigned) > 0:
                    ckt = assigned[0]
                    panel_val = self.get_model_param_value(ckt, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM)
                    cnum_val = self.get_model_param_value(ckt, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER)
        except Exception:
            pass

        if not panel_val:
            panel_val = self.get_model_param_value(device, "Panel")
        if not cnum_val:
            cnum_val = self.get_model_param_value(device, "Circuit Number")

        if panel_val and cnum_val:
            return (str(panel_val), str(cnum_val))
        return None

    def build_detail_lookup(self, detail_items):
        detail_by_component_id = {}
        detail_by_id = {}
        detail_by_circuit = {}
        detail_by_panel = {}

        for ditem in detail_items:
            detail_by_id[str(ditem.Id.IntegerValue)] = ditem

            component_id = self.get_detail_param_value(ditem, SLD_COMPONENT_ID_PARAM)
            if component_id:
                detail_by_component_id[str(component_id).strip()] = ditem

            panel_val = self.get_detail_param_value(ditem, DETAIL_PARAM_CKT_PANEL)
            cnum_val = self.get_detail_param_value(ditem, DETAIL_PARAM_CKT_NUMBER)
            if panel_val and cnum_val:
                detail_by_circuit[(str(panel_val), str(cnum_val))] = ditem

            pname_val = self.get_detail_param_value(ditem, DETAIL_PARAM_PANEL_NAME)
            if pname_val:
                detail_by_panel[str(pname_val)] = ditem
            elif panel_val:
                detail_by_panel[str(panel_val)] = ditem

        return detail_by_component_id, detail_by_id, detail_by_circuit, detail_by_panel

    def get_associated_detail(self, model_elem, model_id, detail_by_component_id, detail_by_id, fallback_key=None,
                              detail_by_circuit=None, detail_by_panel=None):
        detail_elem = None
        component_id = self.get_model_param_value(model_elem, SLD_COMPONENT_ID_PARAM, allow_type_fallback=False)
        if component_id:
            detail_elem = detail_by_component_id.get(str(component_id).strip())

        if not detail_elem and fallback_key:
            if detail_by_circuit and isinstance(fallback_key, tuple):
                detail_elem = detail_by_circuit.get(fallback_key)
            elif detail_by_panel and isinstance(fallback_key, str):
                detail_elem = detail_by_panel.get(fallback_key)

        return detail_elem

    def build_associations(self):
        circuits = self.collect_circuits()
        panels = self.collect_panels()
        devices = self.collect_devices()
        detail_items = self.collect_detail_items()

        detail_by_component_id, detail_by_id, detail_by_circuit, detail_by_panel = self.build_detail_lookup(detail_items)
        associations = []

        for circuit in circuits:
            key = self.get_circuit_key(circuit)
            detail_elem = self.get_associated_detail(circuit, circuit.Id.IntegerValue, detail_by_component_id, detail_by_id,
                                                     fallback_key=key, detail_by_circuit=detail_by_circuit)
            assoc = SyncAssociation(circuit, detail_elem, "circuit", key=key)
            associations.append(assoc)

        for panel in panels:
            key = self.get_panel_key(panel)
            detail_elem = self.get_associated_detail(panel, panel.Id.IntegerValue, detail_by_component_id, detail_by_id,
                                                     fallback_key=key, detail_by_panel=detail_by_panel)
            assoc = SyncAssociation(panel, detail_elem, "panel", key=key)
            associations.append(assoc)

        for device in devices:
            key = self.get_device_circuit_key(device)
            detail_elem = self.get_associated_detail(device, device.Id.IntegerValue, detail_by_component_id, detail_by_id,
                                                     fallback_key=key, detail_by_circuit=detail_by_circuit)
            assoc = SyncAssociation(device, detail_elem, "device", key=key)
            associations.append(assoc)

        for assoc in associations:
            assoc.status = self.compute_status(assoc)

        return associations

    # ------------------------------------------------------------------
    # Status + syncing
    # ------------------------------------------------------------------
    def compute_status(self, assoc):
        if not assoc.detail_elem:
            return "missing"

        if self.is_outdated(assoc):
            return "outdated"
        return "linked"

    def is_outdated(self, assoc):
        value_map = self.get_value_map(assoc)

        for detail_param, model_param in value_map.items():
            model_val = self.get_model_param_value(assoc.model_elem, model_param)
            detail_val = self.get_detail_param_value(assoc.detail_elem, detail_param)
            if self._normalize(model_val) != self._normalize(detail_val):
                return True

        return False

    def get_value_map(self, assoc):
        if assoc.kind == "circuit":
            return CIRCUIT_VALUE_MAP
        if assoc.kind == "panel":
            return PANEL_VALUE_MAP
        return DEVICE_VALUE_MAP

    def compare_values(self, assoc):
        results = []
        value_map = self.get_value_map(assoc)
        for detail_param, model_param in value_map.items():
            model_val = self.get_model_param_value(assoc.model_elem, model_param)
            detail_val = self.get_detail_param_value(assoc.detail_elem, detail_param) if assoc.detail_elem else None
            results.append({
                "param": detail_param,
                "model": model_val,
                "detail": detail_val,
                "match": self._normalize(model_val) == self._normalize(detail_val)
            })
        return results

    def sync_associations(self, associations):
        updated = 0

        for assoc in associations:
            if not assoc.detail_elem:
                continue

            if assoc.kind == "circuit":
                updated += self._sync_circuit(assoc)
            elif assoc.kind == "panel":
                updated += self._sync_panel(assoc)
            else:
                updated += self._sync_device(assoc)

        return updated

    def get_link_warnings(self, associations):
        warnings = []
        for assoc in associations:
            if not assoc.detail_elem:
                warnings.append("{} {} (Id {}) missing detail item".format(
                    assoc.kind.capitalize(),
                    assoc.model_elem.Name if hasattr(assoc.model_elem, "Name") else "Element",
                    assoc.model_elem.Id.IntegerValue))
                continue

            model_val = self.get_model_param_value(assoc.model_elem, SLD_COMPONENT_ID_PARAM, allow_type_fallback=False)
            detail_val = self.get_detail_param_value(assoc.detail_elem, SLD_COMPONENT_ID_PARAM)
            if not model_val and not detail_val:
                warnings.append("Missing SLD_Component ID_CED for model {} and detail {} ({} Id {}, detail Id {})".format(
                    assoc.kind,
                    assoc.detail_elem.Name if hasattr(assoc.detail_elem, "Name") else "Detail Item",
                    assoc.kind.capitalize(),
                    assoc.model_elem.Id.IntegerValue,
                    assoc.detail_elem.Id.IntegerValue))
        return warnings

    def _sync_component_id(self, assoc):
        model_val = self.get_model_param_value(assoc.model_elem, SLD_COMPONENT_ID_PARAM, allow_type_fallback=False)
        detail_val = self.get_detail_param_value(assoc.detail_elem, SLD_COMPONENT_ID_PARAM)
        if model_val:
            self.set_param_value(assoc.detail_elem, SLD_COMPONENT_ID_PARAM, str(model_val))
        elif detail_val:
            self.set_param_value(assoc.model_elem, SLD_COMPONENT_ID_PARAM, str(detail_val))
        else:
            default_id = str(assoc.model_elem.Id.IntegerValue)
            self.set_param_value(assoc.model_elem, SLD_COMPONENT_ID_PARAM, default_id)
            self.set_param_value(assoc.detail_elem, SLD_COMPONENT_ID_PARAM, default_id)

    def _sync_circuit(self, assoc):
        for detail_param, model_param in CIRCUIT_VALUE_MAP.items():
            value = self.get_model_param_value(assoc.model_elem, model_param)
            self.set_param_value(assoc.detail_elem, detail_param, value)

        self.set_param_value(assoc.detail_elem, DETAIL_PARAM_SC_CIRCUIT_ID, str(assoc.model_elem.Id.IntegerValue))
        try:
            panel_elem = assoc.model_elem.BaseEquipment
            if panel_elem:
                self.set_param_value(assoc.detail_elem, DETAIL_PARAM_SC_PANEL_ID, str(panel_elem.Id.IntegerValue))
        except Exception:
            pass

        self._sync_component_id(assoc)
        return 1

    def _sync_panel(self, assoc):
        for detail_param, model_param in PANEL_VALUE_MAP.items():
            value = self.get_model_param_value(assoc.model_elem, model_param)
            self.set_param_value(assoc.detail_elem, detail_param, value)

        self.set_param_value(assoc.detail_elem, DETAIL_PARAM_SC_PANEL_ID, str(assoc.model_elem.Id.IntegerValue))
        self._sync_component_id(assoc)
        return 1

    def _sync_device(self, assoc):
        for detail_param, model_param in DEVICE_VALUE_MAP.items():
            value = self.get_model_param_value(assoc.model_elem, model_param)
            self.set_param_value(assoc.detail_elem, detail_param, value)

        self._sync_component_id(assoc)
        return 1

    # ------------------------------------------------------------------
    # Detail item creation
    # ------------------------------------------------------------------
    def create_detail_items(self, associations, detail_symbol, view, points, tag_symbol=None):
        created = []
        if not detail_symbol or not view or not points:
            return created

        if not detail_symbol.IsActive:
            detail_symbol.Activate()

        for index, assoc in enumerate(associations):
            location = points[index]
            detail_item = self.doc.Create.NewFamilyInstance(location, detail_symbol, view)
            assoc.detail_elem = detail_item
            created.append(assoc)

            if tag_symbol:
                try:
                    tag_point = location
                    tag = DB.IndependentTag.Create(self.doc, view.Id,
                                                   DB.Reference(detail_item),
                                                   False,
                                                   DB.TagMode.TM_ADDBY_CATEGORY,
                                                   DB.TagOrientation.Horizontal,
                                                   tag_point)
                    tag.ChangeTypeId(tag_symbol.Id)
                except Exception:
                    pass

        return created

    # ------------------------------------------------------------------
    # Detail item family helpers
    # ------------------------------------------------------------------
    def collect_detail_symbols(self):
        symbols = DB.FilteredElementCollector(self.doc) \
            .OfCategory(DB.BuiltInCategory.OST_DetailComponents) \
            .WhereElementIsElementType() \
            .ToElements()
        return [sym for sym in symbols
                if isinstance(sym, DB.FamilySymbol)
                and sym.Family
                and sym.Family.Name
                and sym.Family.Name.startswith("SLD-")]

    def collect_tag_symbols(self):
        symbols = DB.FilteredElementCollector(self.doc) \
            .OfCategory(DB.BuiltInCategory.OST_DetailComponentTags) \
            .WhereElementIsElementType() \
            .ToElements()
        return [sym for sym in symbols
                if isinstance(sym, DB.FamilySymbol)
                and sym.Family
                and sym.Family.Name
                and sym.Family.Name.startswith("DI-Tag_SLD")]
