from collections import defaultdict
from pyrevit import DB
from pyrevit.compat import get_elementid_value_func


class CircuitHierarchyReader(object):
    """Loads electrical equipment, circuits, and connected devices."""

    def __init__(self, doc):
        self.doc = doc
        self._hierarchy = []
        self._sorted = []
        self.selected_circuits = []
        self.refresh("System Tree")

    @property
    def hierarchy(self):
        return self._hierarchy

    def refresh(self, sort_mode):
        equipment_lookup = defaultdict(list)
        get_id_val = get_elementid_value_func()

        for circuit in DB.FilteredElementCollector(self.doc)\
                .OfClass(DB.Electrical.ElectricalSystem)\
                .WhereElementIsNotElementType():
            base_equipment = circuit.BaseEquipment
            if not base_equipment:
                continue

            panel_name = getattr(base_equipment, 'Name', None) or "<No Equipment>"
            circuit_label = self._build_circuit_label(circuit, get_id_val)
            devices = self._get_connected_devices(circuit, get_id_val)
            start_slot = getattr(circuit, 'StartSlot', 0)

            equipment_lookup[panel_name].append({
                'label': circuit_label,
                'circuit': circuit,
                'devices': devices,
                'panel': panel_name,
                'start_slot': start_slot,
            })

        panel_names = sorted(equipment_lookup.keys()) if sort_mode == "Alphabetical" else self._sort_system_tree(equipment_lookup)

        hierarchy = []
        for panel_name in panel_names:
            if sort_mode == "Alphabetical":
                circuit_nodes = sorted(equipment_lookup[panel_name], key=lambda c: c['label'])
            else:
                circuit_nodes = sorted(equipment_lookup[panel_name], key=lambda c: (c.get('start_slot', 0), c['label']))
            hierarchy.append({
                'label': panel_name,
                'panel': panel_name,
                'children': circuit_nodes
            })

        self._hierarchy = hierarchy
        self._sorted = panel_names

    def search(self, term):
        if not term:
            return self._hierarchy

        term_lower = term.lower()
        filtered = []
        for panel in self._hierarchy:
            matching_circuits = []
            for circuit in panel['children']:
                device_matches = [d for d in circuit['devices'] if term_lower in d['label'].lower()]
                if term_lower in circuit['label'].lower() or device_matches:
                    circuit_copy = circuit.copy()
                    circuit_copy['devices'] = device_matches if device_matches else circuit['devices']
                    matching_circuits.append(circuit_copy)
            if term_lower in panel['label'].lower() or matching_circuits:
                filtered.append({
                    'label': panel['label'],
                    'panel': panel['panel'],
                    'children': matching_circuits
                })
        return filtered

    def _sort_system_tree(self, lookup):
        ordered_panels = []
        for panel_name, circuits in lookup.items():
            ordered_panels.append((panel_name, min(c.get('start_slot', 0) for c in circuits)))
        return [name for name, _ in sorted(ordered_panels, key=lambda x: (x[1], x[0]))]

    def _get_connected_devices(self, circuit, get_id_val):
        devices = []
        for el in circuit.Elements:
            category = getattr(el, 'Category', None)
            if not category:
                continue
            label = "[{}] {}".format(get_id_val(el.Id), getattr(el, 'Name', ''))
            devices.append({'label': label, 'element': el})
        return sorted(devices, key=lambda d: d['label'])

    def _build_circuit_label(self, circuit, get_id_val):
        panel = getattr(circuit.BaseEquipment, 'Name', None) or "<No Panel>"
        number = getattr(circuit, 'CircuitNumber', '')
        load = getattr(circuit, 'LoadName', '') or ''
        return "[{}] {}/{} - {}".format(get_id_val(circuit.Id), panel, number, load.strip())
