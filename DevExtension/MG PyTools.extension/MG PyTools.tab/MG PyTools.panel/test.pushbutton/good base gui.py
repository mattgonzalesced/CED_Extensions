from pyrevit import forms
from System.Collections.ObjectModel import ObservableCollection
from System.Dynamic import ExpandoObject

# WPF layout with a single dropdown and a DataGrid
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
EMOCP = 80
MMOCP = 80
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
                {"Element1": "MOCP: " + str(MMOCP), "Element2": "MOCP: " + str(EMOCP)}
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
