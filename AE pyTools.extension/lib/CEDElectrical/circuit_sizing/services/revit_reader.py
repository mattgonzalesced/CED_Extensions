from pyrevit import DB


class CircuitListProvider(object):
    """Builds a flat, panel-aware circuit list for selection UI."""

    def __init__(self, doc):
        self.doc = doc
        self._circuits = []
        self._panels = []
        self._build_list()

    @property
    def panels(self):
        return self._panels

    def _build_list(self):
        circuits = []
        panels = set()
        for circuit in DB.FilteredElementCollector(self.doc)\
                .OfClass(DB.Electrical.ElectricalSystem)\
                .WhereElementIsNotElementType():
            base_equipment = circuit.BaseEquipment
            panel_name = getattr(base_equipment, 'Name', None)
            rating = self._safe_int(getattr(circuit, 'Rating', None)) if circuit.SystemType == DB.Electrical.ElectricalSystemType.PowerCircuit else None
            poles = getattr(circuit, 'PolesNumber', None)
            start_slot = getattr(circuit, 'StartSlot', 0)

            if panel_name:
                panels.add(panel_name)

            circuits.append({
                'id': circuit.Id.IntegerValue,
                'panel': panel_name,
                'label': self._format_label(panel_name, circuit.CircuitNumber, circuit.LoadName, rating, poles),
                'circuit': circuit,
                'sort_key': (0 if panel_name else 1, (panel_name or '').lower(), start_slot, (circuit.CircuitNumber or '').lower())
            })

        self._circuits = sorted(circuits, key=lambda c: c['sort_key'])
        self._panels = sorted(panels)

    def filter(self, search_text='', panel_filter=None):
        term = (search_text or '').lower()
        matches = []
        for entry in self._circuits:
            if panel_filter and panel_filter != 'All Panels':
                if panel_filter == '<No Panel>':
                    if entry['panel']:
                        continue
                elif entry['panel'] != panel_filter:
                    continue
            if term and term not in entry['label'].lower():
                continue
            matches.append(entry)
        return matches

    def _format_label(self, panel, circuit_number, load_name, rating, poles):
        panel_label = panel or '<No Panel>'
        load = (load_name or '').strip()
        rating_str = 'N/A' if rating is None else rating
        poles_str = poles if poles is not None else '?'
        return "{}/{} - {} ({}/{})".format(panel_label, circuit_number, load, rating_str, "{}P".format(poles_str))

    def _safe_int(self, value):
        try:
            return int(round(value, 0)) if value is not None else None
        except Exception:
            return None
