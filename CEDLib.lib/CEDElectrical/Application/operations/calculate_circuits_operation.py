# -*- coding: utf-8 -*-

"""Calculate-circuits application operation."""

from datetime import datetime

from pyrevit import DB, forms, script

from CEDElectrical.Domain import settings_manager
from CEDElectrical.Model.CircuitBranch import CircuitBranch
from CEDElectrical.Model.circuit_settings import CircuitSettings
from Snippets import revit_helpers


def _elid_value(item):
    return revit_helpers.get_elementid_value(item)


def _elid_from_value(value):
    return revit_helpers.elementid_from_value(value)


class CalculateCircuitsOperation(object):
    """Orchestrates calculation, writes, and alert persistence for circuits."""
    key = 'calculate_circuits'

    def __init__(self, repository, writer, alert_store):
        self.repository = repository
        self.writer = writer
        self.alert_store = alert_store
        self.logger = script.get_logger()

    def execute(self, request, doc):
        """Run calculation workflow for target circuits in the active document."""
        param_bootstrap = settings_manager.ensure_electrical_parameters_for_calculate(doc, logger=self.logger)
        status = str((param_bootstrap or {}).get('status') or '').lower()
        if status == 'loaded':
            self.logger.info(
                'Auto-loaded electrical parameters for calculate. updated={} unchanged={} skipped={}'.format(
                    int((param_bootstrap or {}).get('updated') or 0),
                    int((param_bootstrap or {}).get('unchanged') or 0),
                    int((param_bootstrap or {}).get('skipped') or 0),
                )
            )
        elif status == 'failed':
            self.logger.warning(
                'Auto-load electrical parameters before calculate failed: {}'.format(
                    (param_bootstrap or {}).get('reason') or 'unknown'
                )
            )

        settings = settings_manager.load_circuit_settings(doc)
        min_breaker_size_override = request.options.get('min_breaker_size_override')
        if min_breaker_size_override is not None:
            try:
                override_value = int(min_breaker_size_override)
                if override_value > 0:
                    settings = CircuitSettings.from_json(settings.to_json())
                    settings.set('min_breaker_size', override_value)
            except Exception:
                pass
        circuits = self.repository.get_target_circuits(doc, request.circuit_ids)

        circuits, locked_ids, locked_rows = self.repository.partition_locked_elements(doc, circuits, settings)
        if locked_ids:
            summary = self.repository.summarize_locked(doc, locked_ids)
            self.logger.info(
                'Locked elements detected; proceeding with editable set only. circuits={} fixtures={} equipment={} other={}'.format(
                    int(summary.get('circuits') or 0),
                    int(summary.get('fixtures') or 0),
                    int(summary.get('equipment') or 0),
                    int(summary.get('other') or 0),
                )
            )

        if not circuits:
            forms.alert('No editable circuits found to process.')
            return {'status': 'cancelled', 'reason': 'no_circuits'}

        count = len(circuits)
        if count > 1000:
            proceed = forms.alert(
                '{} circuits selected.\n\nThis may take a while.\n\n'.format(count),
                title='Large Selection Warning',
                options=['Continue', 'Cancel']
            )
            if proceed != 'Continue':
                return {'status': 'cancelled', 'reason': 'large_selection_cancel'}

        branches = []
        for circuit in circuits:
            branch = CircuitBranch(circuit, settings=settings)
            if not branch.is_power_circuit or branch.is_space or branch.is_spare:
                continue

            branch.calculate_hot_wire_size()
            branch.calculate_neutral_wire_size()
            branch.calculate_ground_wire_size()
            branch.calculate_isolated_ground_wire_size()
            branch.calculate_conduit_size()
            branches.append(branch)

        if not branches:
            forms.alert('No editable branch circuits found to process.')
            return {'status': 'cancelled', 'reason': 'no_branches'}

        total_fixtures = 0
        total_equipment = 0

        use_existing_group = bool(request.options.get('use_existing_transaction_group', False))
        tg = None
        if not use_existing_group:
            tg = DB.TransactionGroup(doc, 'Calculate Circuits')
            tg.Start()
        tx = DB.Transaction(doc, 'Write Shared Parameters')

        try:
            tx.Start()
            for branch in branches:
                param_values = self._collect_shared_param_values(branch)
                self.writer.write_circuit_parameters(branch.circuit, param_values)
                f_cnt, e_cnt = self.writer.write_connected_elements(branch, param_values, settings, locked_ids)
                total_fixtures += f_cnt
                total_equipment += e_cnt

                alert_payload = self._build_alert_payload(branch)
                if alert_payload is None:
                    self.alert_store.clear_alert_payload(branch.circuit)
                else:
                    self.alert_store.write_alert_payload(branch.circuit, alert_payload)

            self._write_locked_sync_payloads(doc, locked_rows)

            tx.Commit()
            if tg is not None:
                tg.Assimilate()
        except Exception as ex:
            try:
                tx.RollBack()
            except Exception:
                pass
            try:
                if tg is not None:
                    tg.RollBack()
            except Exception:
                pass
            self.logger.error('CalculateCircuitsOperation failed: {}'.format(ex))
            raise

        show_output = bool(request.options.get('show_output', True))
        if show_output:
            self._print_report(branches, total_fixtures, total_equipment, locked_rows)
        runtime_alert_rows = self._collect_runtime_alert_rows(branches)
        return {
            'status': 'ok',
            'updated_circuits': len(branches),
            'updated_fixtures': total_fixtures,
            'updated_equipment': total_equipment,
            'locked_rows': locked_rows,
            'runtime_alert_rows': runtime_alert_rows,
        }

    def _collect_shared_param_values(self, branch):
        """Map branch results into shared-parameter values."""
        neutral_qty = branch.neutral_wire_quantity or 0
        ig_qty = branch.isolated_ground_wire_quantity or 0
        include_neutral = 1 if neutral_qty > 0 else 0
        include_ig = 1 if ig_qty > 0 else 0

        return {
            'CKT_Circuit Type_CEDT': branch.branch_type,
            'CKT_Panel_CEDT': branch.panel,
            'CKT_Circuit Number_CEDT': branch.circuit_number,
            'CKT_Load Name_CEDT': branch.load_name,
            'CKT_Rating_CED': branch.rating,
            'CKT_Frame_CED': branch.frame,
            'CKT_Length_CED': branch.length,
            'CKT_Schedule Notes_CEDT': branch.circuit_notes,
            'Voltage Drop Percentage_CED': branch.voltage_drop_percentage,
            'CKT_Wire Hot Size_CEDT': branch.hot_wire_size,
            'CKT_Number of Wires_CED': branch.number_of_wires,
            'CKT_Number of Sets_CED': branch.number_of_sets,
            'CKT_Wire Hot Quantity_CED': branch.hot_wire_quantity,
            'CKT_Wire Ground Size_CEDT': branch.ground_wire_size,
            'CKT_Wire Ground Quantity_CED': branch.ground_wire_quantity,
            'CKT_Wire Neutral Size_CEDT': branch.neutral_wire_size,
            'CKT_Wire Neutral Quantity_CED': neutral_qty,
            'CKT_Wire Isolated Ground Size_CEDT': branch.isolated_ground_wire_size,
            'CKT_Wire Isolated Ground Quantity_CED': ig_qty,
            'CKT_Include Neutral_CED': include_neutral,
            'CKT_Include Isolated Ground_CED': include_ig,
            'Wire Material_CEDT': branch.wire_material,
            'Wire Temparature Rating_CEDT': branch.wire_temp_rating,
            'Wire Insulation_CEDT': branch.wire_insulation,
            'Conduit Size_CEDT': branch.conduit_size,
            'Conduit Type_CEDT': branch.conduit_type,
            'Conduit Fill Percentage_CED': branch.conduit_fill_percentage,
            'Wire Size_CEDT': branch.get_wire_size_callout(),
            'Conduit and Wire Size_CEDT': branch.get_conduit_and_wire_size(),
            'Circuit Load Current_CED': branch.circuit_load_current,
            'Circuit Ampacity_CED': branch.circuit_base_ampacity,
            'CKT_Length Makeup_CED': branch.wire_length_makeup,
        }

    def _build_alert_payload(self, branch):
        """Build serializable alert payload for persistence."""
        notices = getattr(branch, 'notices', None)
        if not notices or not notices.has_items():
            return None

        existing = self.alert_store.read_alert_payload(branch.circuit) or {}
        existing_hidden = existing.get('hidden_definition_ids') if isinstance(existing, dict) else []
        if not isinstance(existing_hidden, list):
            existing_hidden = []

        items = []
        present_ids = set()
        for definition, severity, group, message in notices.items:
            if definition and not getattr(definition, 'persistent', True):
                continue
            definition_id = definition.GetId() if definition else None
            if definition_id:
                present_ids.add(definition_id)
            items.append({
                'definition_id': definition_id,
                'severity': severity,
                'group': group,
                'message': message,
            })

        hidden_ids = sorted(list(set(existing_hidden).intersection(present_ids)))
        payload = {
            'version': 1,
            'generated_utc': datetime.utcnow().isoformat() + 'Z',
            'circuit': {
                'id': _elid_value(branch.circuit.Id),
                'name': branch.name,
                'panel': branch.panel,
                'number': branch.circuit_number,
            },
            'alerts': items,
            'hidden_definition_ids': hidden_ids,
        }
        return payload

    def _print_report(self, branches, total_fixtures, total_equipment, locked_rows):
        """Print a post-run report to pyRevit output."""
        output = script.get_output()
        try:
            output.show()
        except Exception:
            pass
        output.close_others()
        output.print_md('## Shared Parameters Updated')
        output.print_md('* Circuits updated: **{}**'.format(len(branches)))
        output.print_md('* Electrical Fixtures updated: **{}**'.format(total_fixtures))
        output.print_md('* Electrical Equipment updated: **{}**'.format(total_equipment))

        if locked_rows:
            output.print_md('\n## Skipped Elements')
            output.print_md('The following elements are owned by other users and could not be calculated.')
            table = []
            for row in locked_rows:
                table.append([
                    row.get('circuit', ''),
                    row.get('circuit_owner', '') or '-',
                    row.get('device_owner', '') or '-',
                ])
            output.print_table(table_data=table, columns=['Circuit', 'Circuit Owner', 'Device Owner'])

        label_map = {
            'Overrides': 'Overrides',
            'Calculation': 'Calculation',
            'Design': 'Design',
            'Error': 'Error',
            'Other': 'Other',
        }
        severity_colors = {
            'NONE': None,
            'MEDIUM': '#d9822b',
            'HIGH': '#d9534f',
            'CRITICAL': '#b20000',
        }

        notice_lines = []
        for branch in branches:
            if not getattr(branch, 'notices', None) or not branch.notices.has_items():
                continue
            notice_lines.extend(branch.notices.formatted_lines(label_map, severity_colors))

        if notice_lines:
            output.print_md('\n## Warnings / Errors')
            for line in notice_lines:
                output.print_md(line)
        try:
            output.show()
        except Exception:
            pass

    def _collect_runtime_alert_rows(self, branches):
        rows = []
        for branch in branches:
            notices = getattr(branch, 'notices', None)
            if not notices or not notices.has_items():
                continue
            for definition, severity, group, message in notices.items:
                if definition is not None and getattr(definition, 'persistent', True):
                    continue
                definition_id = ''
                try:
                    definition_id = definition.GetId() if definition else ''
                except Exception:
                    definition_id = ''
                rows.append({
                    'panel': branch.panel or '',
                    'number': branch.circuit_number or '',
                    'load_name': branch.load_name or '',
                    'group': group or 'Other',
                    'definition_id': definition_id or '-',
                    'message': message or '',
                })
        return rows

    def _write_locked_sync_payloads(self, doc, locked_rows):
        for row in list(locked_rows or []):
            try:
                if not bool(row.get('sync_writeback', False)):
                    continue
                circuit_id = int(row.get('circuit_id') or 0)
                if circuit_id <= 0:
                    continue
                circuit = doc.GetElement(_elid_from_value(circuit_id))
                if circuit is None:
                    continue
                payload = self.alert_store.read_alert_payload(circuit) or {}
                if not isinstance(payload, dict):
                    payload = {}
                alerts = payload.get('alerts')
                payload['alerts'] = alerts if isinstance(alerts, list) else []
                hidden = payload.get('hidden_definition_ids')
                payload['hidden_definition_ids'] = hidden if isinstance(hidden, list) else []
                payload['version'] = payload.get('version') or 1
                payload['sync_lock'] = {
                    'blocked': True,
                    'generated_utc': datetime.utcnow().isoformat() + 'Z',
                    'circuit_owner': row.get('circuit_owner') or '',
                    'device_owner': row.get('device_owner') or '',
                }
                self.alert_store.write_alert_payload(circuit, payload)
            except Exception:
                continue

