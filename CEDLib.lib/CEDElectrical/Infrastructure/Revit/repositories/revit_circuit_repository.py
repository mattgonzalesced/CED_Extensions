# -*- coding: utf-8 -*-
"""Revit-backed circuit repository adapter."""

import Autodesk.Revit.DB.Electrical as DBE
from pyrevit import DB
from pyrevit.compat import get_elementid_value_func, get_elementid_from_value_func


_get_elid_value = get_elementid_value_func()
_get_elid_from_value = get_elementid_from_value_func()


def _elid_value(item):
    try:
        return int(_get_elid_value(item))
    except Exception:
        return int(getattr(item, "IntegerValue", 0))


def _elid_from_value(value):
    return _get_elid_from_value(int(value))


class RevitCircuitRepository(object):
    """Loads circuit targets and lock metadata from Revit."""

    def get_target_circuits(self, doc, circuit_ids=None):
        """Return circuits from explicit ids or all project circuits."""
        circuit_ids = list(circuit_ids or [])
        if circuit_ids:
            circuits = []
            for raw_id in circuit_ids:
                try:
                    el = doc.GetElement(_elid_from_value(raw_id))
                except Exception:
                    el = None
                if isinstance(el, DBE.ElectricalSystem):
                    circuits.append(el)
            return circuits

        return list(
            DB.FilteredElementCollector(doc)
            .OfClass(DBE.ElectricalSystem)
            .WhereElementIsNotElementType()
            .ToElements()
        )

    def partition_locked_elements(self, doc, circuits, settings, collect_all_device_owners=True):
        """Split circuits into editable and locked subsets."""
        if not getattr(doc, 'IsWorkshared', False):
            return circuits, set(), []

        locked_ids = set()
        unlocked_circuits = []
        locked_records = {}

        def _is_locked(eid):
            try:
                status = DB.WorksharingUtils.GetCheckoutStatus(doc, eid)
                return status == DB.CheckoutStatus.OwnedByOtherUser
            except Exception:
                return False

        def _owner_name(eid):
            try:
                info = DB.WorksharingUtils.GetWorksharingTooltipInfo(doc, eid)
                return info.Owner
            except Exception:
                return None

        def _circuit_label(circuit):
            panel = getattr(circuit.BaseEquipment, 'Name', '') if circuit.BaseEquipment else ''
            number = getattr(circuit, 'CircuitNumber', '') or ''
            return '{}-{}'.format(panel, number)

        def _ensure_record(circuit):
            key = _elid_value(circuit.Id)
            if key not in locked_records:
                locked_records[key] = {
                    'circuit_id': key,
                    'circuit': _circuit_label(circuit),
                    'load_name': getattr(circuit, 'LoadName', '') or '',
                    'circuit_owner': '',
                    'circuit_locked': False,
                    'device_owners': set(),
                }
            return locked_records[key]

        write_fixtures = getattr(settings, 'write_fixture_results', False)
        write_equipment = getattr(settings, 'write_equipment_results', False)

        for circuit in circuits:
            locked_for_writeback = False
            if _is_locked(circuit.Id):
                locked_ids.add(circuit.Id)
                record = _ensure_record(circuit)
                record['circuit_locked'] = True
                record['circuit_owner'] = _owner_name(circuit.Id) or ''
                continue

            if write_equipment or write_fixtures:
                for el in circuit.Elements:
                    if not isinstance(el, DB.FamilyInstance):
                        continue
                    cat = el.Category
                    if not cat:
                        continue
                    cat_id = cat.Id
                    is_fixture = cat_id == DB.ElementId(DB.BuiltInCategory.OST_ElectricalFixtures)
                    is_equipment = cat_id == DB.ElementId(DB.BuiltInCategory.OST_ElectricalEquipment)

                    if is_fixture and not write_fixtures:
                        continue
                    if is_equipment and not write_equipment:
                        continue

                    if _is_locked(el.Id):
                        locked_ids.add(el.Id)
                        rec = _ensure_record(circuit)
                        owner = _owner_name(el.Id)
                        if owner:
                            rec['device_owners'].add(owner)
                        locked_for_writeback = True
                        if not collect_all_device_owners:
                            break

            if locked_for_writeback:
                locked_ids.add(circuit.Id)
                _ensure_record(circuit)
                continue

            unlocked_circuits.append(circuit)

        locked_rows = []
        for rec in locked_records.values():
            locked_rows.append({
                'circuit_id': rec.get('circuit_id') or 0,
                'circuit': rec['circuit'],
                'load_name': rec.get('load_name') or '',
                'circuit_owner': rec.get('circuit_owner') or '',
                'device_owner': ', '.join(sorted(rec['device_owners'])) if rec['device_owners'] else '',
                'circuit_locked': bool(rec.get('circuit_locked', False)),
                'sync_writeback': not bool(rec.get('circuit_locked', False)),
            })

        return unlocked_circuits, locked_ids, locked_rows

    def summarize_locked(self, doc, locked_ids):
        """Return lock summary counts by element type."""
        summary = {'circuits': 0, 'fixtures': 0, 'equipment': 0, 'other': 0}
        for eid in locked_ids:
            el = doc.GetElement(eid)
            if isinstance(el, DBE.ElectricalSystem):
                summary['circuits'] += 1
                continue
            if isinstance(el, DB.FamilyInstance):
                cat = el.Category
                if cat:
                    cid = cat.Id
                    if cid == DB.ElementId(DB.BuiltInCategory.OST_ElectricalFixtures):
                        summary['fixtures'] += 1
                        continue
                    if cid == DB.ElementId(DB.BuiltInCategory.OST_ElectricalEquipment):
                        summary['equipment'] += 1
                        continue
            summary['other'] += 1
        return summary
