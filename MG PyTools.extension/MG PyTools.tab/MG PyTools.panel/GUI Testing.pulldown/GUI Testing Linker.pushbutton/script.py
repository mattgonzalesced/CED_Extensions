# -*- coding: utf-8 -*-
__title__ = "Parameter Linker"
from pyrevit import revit, DB, forms
from System.Collections.ObjectModel import ObservableCollection
from System.Dynamic import ExpandoObject

# WPF layout with a dropdown and DataGrid
layout = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Parameter Set Selector"
        Height="600" Width="400"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize">
    <Window.Background>
        <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
            <GradientStop Color="#FF1E88E5" Offset="0.0" />
            <GradientStop Color="#FF64B5F6" Offset="1.0" />
        </LinearGradientBrush>
    </Window.Background>
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>

        <!-- Title -->
        <TextBlock Grid.Row="0" Text="Select a Parameter Set" FontSize="20" FontWeight="Bold" Foreground="White" 
                   HorizontalAlignment="Center" Margin="0,0,0,10" />

        <!-- Dropdown -->
        <StackPanel Grid.Row="1" VerticalAlignment="Top" Margin="20">
            <TextBlock Text="Parameter Sets:" FontSize="16" Foreground="White" Margin="0,0,0,10" />
            <ComboBox Name="dropdown" Width="200" Margin="0,0,0,10" />
            <TextBlock Text="Details in Set:" FontSize="16" Foreground="White" Margin="0,10,0,5" />
            <DataGrid Name="detailsGrid" AutoGenerateColumns="False" Height="300" Background="White" IsReadOnly="True">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Element 1" Binding="{Binding Element1}" Width="*" />
                    <DataGridTextColumn Header="Element 2" Binding="{Binding Element2}" Width="*" />
                </DataGrid.Columns>
            </DataGrid>
        </StackPanel>

        <!-- Confirm Button -->
        <Button Grid.Row="2" Content="Confirm" FontSize="16" FontWeight="Bold"
                HorizontalAlignment="Center" Width="120" Background="#FF64B5F6" Foreground="White"
                Margin="0,10" />
    </Grid>
</Window>
"""

# Python code to handle data and dropdown
class ParameterSetSelector(forms.WPFWindow):
    def __init__(self):
        super(ParameterSetSelector, self).__init__(layout, literal_string=True)

        # Dropdown options
        self.dropdown_options = ["One Line Diagram", "Load Calculation", "Voltage Mech"]

        # Data for the dropdown options
        self.data_sets = {
            "One Line Diagram": [
                {"Element1": "Panel Name", "Element2": "Name"},
                {"Element1": "Voltage", "Element2": "Voltage"},
                {"Element1": "ISC", "Element2": "ISC"},
                {"Element1": "VD", "Element2": "VD"},
            ],
            "Load Calculation": [
                {"Element1": "Load Name", "Element2": "Value"},
                {"Element1": "Type", "Element2": "Category"},
                {"Element1": "Demand", "Element2": "Total Load"},
            ],
            "Voltage Mech": [
                {"Element1": "MOCP", "Element2": "EMOCP"}
            ]
        }

        # Populate the dropdown
        self.dropdown.ItemsSource = self.dropdown_options

        # Event handler for dropdown selection
        self.dropdown.SelectionChanged += self.on_selection_changed

        # Initialize DataGrid
        self.detailsGrid.ItemsSource = ObservableCollection[ExpandoObject]()

    def on_selection_changed(self, sender, args):
        # Get selected value from the dropdown
        selected = self.dropdown.SelectedItem

        # Check if the selected option has data
        if selected and selected in self.data_sets:
            # Populate DataGrid with data for the selected option
            data = self.data_sets[selected]
            converted_data = ObservableCollection[ExpandoObject]()
            for row in data:
                expando = ExpandoObject()
                for k, v in row.items():
                    setattr(expando, k, v)
                converted_data.Add(expando)
            self.detailsGrid.ItemsSource = converted_data
        else:
            # Clear DataGrid if no data is available
            self.detailsGrid.ItemsSource = ObservableCollection[ExpandoObject]()

# Show the custom form
try:
    selector = ParameterSetSelector()  # Create the parameter set selector
    selector.show_dialog()  # Open the window
except Exception as e:
    # In case of critical error, show the error message
    forms.alert("Critical error: {0}".format(str(e)), title="Error")

# Function to apply mappings in Revit
def apply_parameter_mappings():
    """Apply mappings to elements in the Revit document."""
    elements = DB.FilteredElementCollector(revit.doc).WhereElementIsNotElementType()
    for elem in elements:
        for param_a, param_b in parameter_set.get_mappings():
            param_a_obj = elem.LookupParameter(param_a.name)
            param_b_obj = elem.LookupParameter(param_b.name)

            if param_a_obj and param_b_obj:
                if not param_a.is_read_only and not param_b.is_read_only:
                    try:
                        # Example: Copy value from param_a to param_b
                        if param_a.storage_type == "Double":
                            param_b_obj.Set(param_a_obj.AsDouble())
                        elif param_a.storage_type == "Integer":
                            param_b_obj.Set(param_a_obj.AsInteger())
                        elif param_a.storage_type == "String":
                            param_b_obj.Set(param_a_obj.AsString())
                        print("Mapped {0} to {1} for element {2}".format(param_a.name, param_b.name, elem.Id))
                    except Exception as e:
                        print("Error mapping {0} to {1}: {2}".format(param_a.name, param_b.name, e))

# Apply parameter mappings
apply_parameter_mappings()

# Display mappings in the console
for param_a, param_b in parameter_set.get_mappings():
    print("Mapping:")
    print("  Element A - Parameter:", param_a.to_dict())
    print("  Element B - Parameter:", param_b.to_dict())
