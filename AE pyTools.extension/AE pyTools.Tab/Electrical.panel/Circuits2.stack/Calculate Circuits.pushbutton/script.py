# -*- coding: utf-8 -*-
"""Modaless circuit browser and calculator for electrical systems."""
from pyrevit import DB, forms, revit, script
import System.Windows.Media as Media

from CEDElectrical.circuit_sizing.domain.circuit_evaluator import CircuitEvaluator
from CEDElectrical.circuit_sizing.services.revit_reader import RevitCircuitReader
from CEDElectrical.circuit_sizing.services.revit_writer import RevitCircuitWriter
from CEDElectrical.circuit_sizing.models.circuit_branch import CircuitSettings
from Snippets import _elecutils as eu

app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
logger = script.get_logger()


class CircuitListItem:
    """Lightweight item used to visualize and refresh circuit data."""

    def __init__(self, circuit, settings):
        self.circuit = circuit
        self.settings = settings
        self.model = RevitCircuitReader(circuit, settings=settings).to_model()
        self.DisplayName = "{} - {}".format(self.model.panel, self.model.circuit_number)
        self.ColorBrush = self._color_for_branch(self.model.branch_type)

    def refresh(self):
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

    XAML = r"""
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Title="Circuit Browser" Height="480" Width="760" Background="White">
        <Grid Margin="10">
            <Grid.RowDefinitions>
                <RowDefinition Height="*" />
                <RowDefinition Height="Auto" />
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="2*" />
                <ColumnDefinition Width="3*" />
            </Grid.ColumnDefinitions>

            <ListBox x:Name="CircuitList" Grid.Row="0" Grid.Column="0" SelectionChanged="OnCircuitSelected" SelectionMode="Extended">
                <ListBox.ItemTemplate>
                    <DataTemplate>
                        <TextBlock Text="{Binding DisplayName}" Foreground="{Binding ColorBrush}" />
                    </DataTemplate>
                </ListBox.ItemTemplate>
            </ListBox>

            <StackPanel Grid.Row="0" Grid.Column="1" Margin="10,0,0,0">
                <TextBlock Text="Circuit Details" FontWeight="Bold" FontSize="14" />
                <TextBlock Text="Panel:" FontWeight="Bold" />
                <TextBlock x:Name="PanelText" Margin="0,0,0,8" />
                <TextBlock Text="Number:" FontWeight="Bold" />
                <TextBlock x:Name="NumberText" Margin="0,0,0,8" />
                <TextBlock Text="Type:" FontWeight="Bold" />
                <TextBlock x:Name="TypeText" Margin="0,0,0,8" />
                <TextBlock Text="Load (A):" FontWeight="Bold" />
                <TextBlock x:Name="LoadText" Margin="0,0,0,8" />
                <TextBlock Text="Length (ft):" FontWeight="Bold" />
                <TextBlock x:Name="LengthText" Margin="0,0,0,8" />
            </StackPanel>

            <StackPanel Grid.Row="1" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                <Button Content="Calculate Selected" Width="160" Margin="0,0,10,0" Click="OnCalculateClicked" />
                <Button Content="Close" Width="80" Click="OnCloseClicked" />
            </StackPanel>
        </Grid>
    </Window>
    """

    def __init__(self, doc, circuits, settings):
        super(CircuitSelectionWindow, self).__init__(self.XAML)
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
                logger.error("Failed to evaluate circuit {}: {}".format(item.circuit.Id, exc))

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
            logger.error("Transaction failed: {}".format(exc))
            return

        output = script.get_output()
        output.close_others()
        output.print_md("## ✅ Shared Parameters Updated")
        output.print_md("* Circuits updated: **{}**".format(len(evaluated)))
        output.print_md("* Electrical Fixtures updated: **{}**".format(total_fixtures))
        output.print_md("* Electrical Equipment updated: **{}**".format(total_equipment))


def collect_circuits():
    selection = revit.get_selection()
    circuits = []
    if selection:
        circuits = eu.get_circuits_from_selection(selection)
    if not circuits:
        collector = DB.FilteredElementCollector(doc).OfClass(DB.Electrical.ElectricalSystem)
        circuits = [c for c in collector if c.SystemType == DB.Electrical.ElectricalSystemType.PowerCircuit]
    return circuits


def main():
    settings = CircuitSettings()
    circuits = collect_circuits()
    if not circuits:
        forms.alert("No electrical circuits found for display.", title="Circuit Browser")
        return
    window = CircuitSelectionWindow(doc, circuits, settings)
    window.show(modal=False)


if __name__ == "__main__":
    main()
