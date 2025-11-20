# -*- coding: utf-8 -*-
"""Modeless circuit browser window and helpers."""
import os

from pyrevit import DB, forms, revit, script
import System.Windows.Media as Media

from CEDElectrical.circuit_sizing.domain.circuit_evaluator import CircuitEvaluator
from CEDElectrical.circuit_sizing.services.revit_reader import RevitCircuitReader
from CEDElectrical.circuit_sizing.services.revit_writer import RevitCircuitWriter

LOGGER = script.get_logger()
XAML_PATH = os.path.join(os.path.dirname(__file__), "circuit_browser.xaml")


class CircuitListItem(object):
    """Lightweight item used to visualize and refresh circuit data."""

    def __init__(self, circuit, settings):
        self.circuit = circuit
        self.settings = settings
        self.model = RevitCircuitReader(circuit, settings=settings).to_model()
        self.DisplayName = "{} - {}".format(self.model.panel, self.model.circuit_number)
        self.ColorBrush = self._color_for_branch(self.model.branch_type)

    def refresh(self):
        """Reload the model and display values for this circuit."""
        self.model = RevitCircuitReader(self.circuit, settings=self.settings).to_model()
        self.DisplayName = "{} - {}".format(self.model.panel, self.model.circuit_number)
        self.ColorBrush = self._color_for_branch(self.model.branch_type)
        return self

    def _color_for_branch(self, branch_type):
        mapping = {
            "FEEDER": Media.Colors.DarkOrange,
            "BRANCH": Media.Colors.DarkGreen,
            "TRANSFORMER_PRIMARY": Media.Colors.MediumPurple,
            "TRANSFORMER_SECONDARY": Media.Colors.SteelBlue,
            "SPACE": Media.Colors.Gray,
            "SPARE": Media.Colors.LightGray,
        }
        return Media.SolidColorBrush(mapping.get(branch_type, Media.Colors.Black))


class CircuitSelectionWindow(forms.WPFWindow):
    """Modeless list of electrical systems with inline calculation triggers."""

    def __init__(self, doc, circuits, settings):
        super(CircuitSelectionWindow, self).__init__(XAML_PATH)
        self.doc = doc
        self.settings = settings
        self.items = [CircuitListItem(circuit, settings) for circuit in circuits]
        self.CircuitList.ItemsSource = self.items
        if self.items:
            self.CircuitList.SelectedIndex = 0
            self._populate_details(self.items[0])

    def OnCircuitSelected(self, sender, args):
        item = self.CircuitList.SelectedItem
        self._populate_details(item)

    def OnCloseClicked(self, sender, args):
        self.Close()

    def OnCalculateClicked(self, sender, args):
        selected = list(self.CircuitList.SelectedItems) if self.CircuitList.SelectedItems else []
        if not selected:
            selected = self.items
        self._calculate_and_write(selected)

    def _populate_details(self, item):
        if not item:
            self.PanelText.Text = ""
            self.NumberText.Text = ""
            self.TypeText.Text = ""
            self.LoadText.Text = ""
            self.LengthText.Text = ""
            return
        model = item.model
        self.PanelText.Text = str(model.panel)
        self.NumberText.Text = str(model.circuit_number)
        self.TypeText.Text = model.classification
        self.LoadText.Text = "{}".format(model.circuit_load_current or model.apparent_current or "")
        self.LengthText.Text = "{}".format(model.length or "")

    def _calculate_and_write(self, selected_items):
        evaluated = []
        for item in selected_items:
            try:
                refreshed = item.refresh()
                calc_result = CircuitEvaluator.evaluate(refreshed.model)
                evaluated.append((item.circuit, refreshed.model, calc_result))
            except Exception as exc:
                LOGGER.error("Failed to evaluate circuit {}: {}".format(item.circuit.Id, exc))

        if not evaluated:
            forms.alert("No circuits evaluated.", title="Circuit Browser")
            return

        tg = DB.TransactionGroup(self.doc, "Calculate Circuits")
        tg.Start()
        t = DB.Transaction(self.doc, "Write Shared Parameters")
        total_fixtures = 0
        total_equipment = 0

        try:
            t.Start()
            for circuit, model, calc_result in evaluated:
                param_values = RevitCircuitWriter.collect_param_values(model, calc_result)
                RevitCircuitWriter.update_circuit_parameters(circuit, param_values)
                fixtures, equipment = RevitCircuitWriter.update_connected_elements(circuit, param_values)
                total_fixtures += fixtures
                total_equipment += equipment
            t.Commit()
            tg.Assimilate()
        except Exception as exc:
            t.RollBack()
            tg.RollBack()
            LOGGER.error("Transaction failed: {}".format(exc))
            return

        output = script.get_output()
        output.close_others()
        output.print_md("## ✅ Shared Parameters Updated")
        output.print_md("* Circuits updated: **{}**".format(len(evaluated)))
        output.print_md("* Electrical Fixtures updated: **{}**".format(total_fixtures))
        output.print_md("* Electrical Equipment updated: **{}**".format(total_equipment))
