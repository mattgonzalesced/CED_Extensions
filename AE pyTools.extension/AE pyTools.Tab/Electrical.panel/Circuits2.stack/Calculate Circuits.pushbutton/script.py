# -*- coding: utf-8 -*-
from pyrevit import DB, forms, revit, script

from CEDElectrical.Model.CircuitBranch import *
from CEDElectrical.Domain import settings_manager
from Snippets import _elecutils as eu

app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
logger = script.get_logger()


# -------------------------------------------------------------------------
# Collects parameter values from a CircuitBranch object
# -------------------------------------------------------------------------
def collect_shared_param_values(branch):
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
        'CKT_Wire Neutral Quantity_CED': branch.neutral_wire_quantity,
        'CKT_Wire Isolated Ground Size_CEDT': branch.isolated_ground_wire_size,
        'CKT_Wire Isolated Ground Quantity_CED': branch.isolated_ground_wire_quantity,
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
        'CKT_Length Makeup_CED': branch.wire_length_makeup
    }


# -------------------------------------------------------------------------
# Write shared parameters to the electrical circuit
# -------------------------------------------------------------------------
def update_circuit_parameters(circuit, param_values):
    for param_name, value in param_values.items():
        param = circuit.LookupParameter(param_name)
        if not param:
            logger.debug("⚠️ Did not find parameter '{}' on circuit {}".format(param_name, circuit.Id))
            continue

        # --------------------------
        # BLANKING OUT NULL VALUES
        # --------------------------
        if value is None:
            try:
                st = param.StorageType
                if st == DB.StorageType.String:
                    param.Set("")
                elif st == DB.StorageType.Integer:
                    # For Yes/No or integer params, use 0
                    param.Set(0)
                elif st == DB.StorageType.Double:
                    # For numeric params, use 0.0
                    param.Set(0.0)
                elif st == DB.StorageType.ElementId:
                    param.Set(DB.ElementId.InvalidElementId)
                logger.debug("🧹 Cleared '{}' on circuit {}".format(param_name, circuit.Id))
            except Exception as e:
                logger.debug("❌ Failed to blank '{}' on circuit {}: {}".format(param_name, circuit.Id, e))
            continue

        # --------------------------
        # WRITING VALID VALUES
        # --------------------------
        try:
            st = param.StorageType
            if st == DB.StorageType.String:
                param.Set(str(value))
            elif st == DB.StorageType.Integer:
                param.Set(int(value))
            elif st == DB.StorageType.Double:
                param.Set(float(value))
            elif st == DB.StorageType.ElementId:
                # user should pass an ElementId or None
                if isinstance(value, DB.ElementId):
                    param.Set(value)
                else:
                    param.Set(DB.ElementId.InvalidElementId)
        except Exception as e:
            logger.debug("❌ Failed to write '{}' to circuit {}: {}".format(param_name, circuit.Id, e))


# -------------------------------------------------------------------------
# Write shared parameters to connected family instances
# -------------------------------------------------------------------------
def update_connected_elements(branch, param_values, settings, locked_ids=None):
    circuit = branch.circuit
    fixture_count = 0
    equipment_count = 0
    locked_ids = locked_ids or set()

    write_fixtures = getattr(settings, 'write_fixture_results', False)
    write_equipment = getattr(settings, 'write_equipment_results', False)
    if not (write_fixtures or write_equipment):
        return fixture_count, equipment_count

    for el in circuit.Elements:
        if not isinstance(el, DB.FamilyInstance):
            continue

        cat = el.Category
        if not cat:
            continue

        cat_id = cat.Id
        is_fixture = cat_id == DB.ElementId(DB.BuiltInCategory.OST_ElectricalFixtures)
        is_equipment = cat_id == DB.ElementId(DB.BuiltInCategory.OST_ElectricalEquipment)

        if not (is_fixture or is_equipment):
            continue

        if el.Id in locked_ids:
            continue

        if is_fixture and not write_fixtures:
            continue
        if is_equipment and not write_equipment:
            continue

        # Write all parameters
        for param_name, value in param_values.items():
            if value is None:
                continue
            param = el.LookupParameter(param_name)
            if not param:
                continue
            try:
                if param.StorageType == DB.StorageType.String:
                    param.Set(str(value))
                elif param.StorageType == DB.StorageType.Integer:
                    param.Set(int(value))
                elif param.StorageType == DB.StorageType.Double:
                    param.Set(float(value))
            except Exception as e:
                logger.debug("❌ Failed to write '{}' to element {}: {}".format(param_name, el.Id, e))

        if is_fixture:
            fixture_count += 1
        elif is_equipment:
            equipment_count += 1

    return fixture_count, equipment_count


def _partition_locked_elements(doc, circuits, settings):
    """Separate locked elements and return unlocked circuits + locked ids."""
    if not getattr(doc, "IsWorkshared", False):
        return circuits, set()

    locked_ids = set()
    unlocked_circuits = []

    def _is_locked(eid):
        try:
            status = DB.WorksharingUtils.GetCheckoutStatus(doc, eid)
            return status == DB.CheckoutStatus.OwnedByOther
        except Exception:
            return False

    downstream_ids = set()
    write_fixtures = getattr(settings, 'write_fixture_results', False)
    write_equipment = getattr(settings, 'write_equipment_results', False)

    for circuit in circuits:
        if _is_locked(circuit.Id):
            locked_ids.add(circuit.Id)
            continue
        unlocked_circuits.append(circuit)

        if not (write_equipment or write_fixtures):
            continue

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
            downstream_ids.add(el.Id)

    for eid in downstream_ids:
        if _is_locked(eid):
            locked_ids.add(eid)

    return unlocked_circuits, locked_ids


def _summarize_locked(doc, locked_ids):
    summary = {'circuits': 0, 'fixtures': 0, 'equipment': 0, 'other': 0}
    for eid in locked_ids:
        el = doc.GetElement(eid)
        if isinstance(el, DB.Electrical.ElectricalSystem):
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


# -------------------------------------------------------------------------
# Main Execution
# -------------------------------------------------------------------------
def main():
    selection = revit.get_selection()
    test_circuits = []
    settings = settings_manager.load_circuit_settings(doc)
    if not selection:
        test_circuits = eu.pick_circuits_from_list(doc, select_multiple=True)
    else:
        for el in selection:
            if isinstance(el, DB.Electrical.ElectricalSystem):
                test_circuits.append(el)
        if not test_circuits:
            test_circuits = eu.pick_circuits_from_list(doc, select_multiple=True)

    test_circuits, locked_ids = _partition_locked_elements(doc, test_circuits, settings)
    if locked_ids:
        summary = _summarize_locked(doc, locked_ids)
        msg_lines = [
            "Some elements are owned by others and will be skipped:",
            "• Circuits: {}".format(summary['circuits']),
        ]
        if settings.write_fixture_results:
            msg_lines.append("• Fixtures: {}".format(summary['fixtures']))
        if settings.write_equipment_results:
            msg_lines.append("• Electrical Equipment: {}".format(summary['equipment']))
        if summary['other']:
            msg_lines.append("• Other: {}".format(summary['other']))
        choice = forms.alert("\n".join(msg_lines), options=["Continue with Unlocked", "Cancel"], ok=False, yes=True, no=True)
        if choice != "Continue with Unlocked":
            script.exit()

    count = len(test_circuits)
    if count > 1000:
        proceed = forms.alert(
            "{} circuits selected.\n\nThis may take a while.\n\n".format(count),
            title="⚠️ Large Selection Warning",
            options=["Continue", "Cancel"]
        )
        if proceed != "Continue":
            script.exit()

    branches = []
    total_fixtures = 0
    total_equipment = 0

    if not test_circuits:
        forms.alert("No editable circuits found to process.")
        return

    # Perform all calculations first
    for circuit in test_circuits:
        branch = CircuitBranch(circuit, settings=settings)
        if not branch.is_power_circuit:
            continue

        branch.calculate_hot_wire_size()
        branch.calculate_neutral_wire_size()
        branch.calculate_ground_wire_size()
        branch.calculate_conduit_size()
        branches.append(branch)

    # Write all parameters in a single transaction
    tg = DB.TransactionGroup(doc, "Calculate Circuits")
    tg.Start()
    t = DB.Transaction(doc, "Write Shared Parameters")
    try:
        t.Start()
        for branch in branches:
            param_values = collect_shared_param_values(branch)
            update_circuit_parameters(branch.circuit, param_values)
            f, e = update_connected_elements(branch, param_values, settings, locked_ids)
            total_fixtures += f
            total_equipment += e
        t.Commit()
        tg.Assimilate()

        output = script.get_output()
        output.close_others()
        output.print_md("## ✅ Shared Parameters Updated")
        output.print_md("* Circuits updated: **{}**".format(len(branches)))
        output.print_md("* Electrical Fixtures updated: **{}**".format(total_fixtures))
        output.print_md("* Electrical Equipment updated: **{}**".format(total_equipment))

    except Exception as e:
        t.RollBack()
        tg.RollBack()
        logger.error("{}❌ Transaction failed: {}".format(branch.name,e))
        return



if __name__ == "__main__":
    main()
