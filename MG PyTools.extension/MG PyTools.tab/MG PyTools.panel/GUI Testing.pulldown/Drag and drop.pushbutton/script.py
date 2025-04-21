import csv
import os
from pyrevit import forms
from System.Collections.ObjectModel import ObservableCollection
from System.Dynamic import ExpandoObject
from System.Windows import DragDrop, DragDropEffects, DataObject

class ParameterSetSelector(forms.WPFWindow):

    def __init__(self):
        # Define the CSV file location to the specified path
        self.CSV_FILE = r"C:\Users\m.gonzales\OneDrive - CoolSys Inc\Desktop\CSV Drag and drop\parameter_set_data.csv"

        # XAML Layout (inline for simplicity)
        layout = """
        <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                Title="Parameter Set Selector"
                Height="400" Width="600"
                WindowStartupLocation="CenterScreen">
            <Grid Margin="10">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="*" />
                    <RowDefinition Height="Auto" />
                </Grid.RowDefinitions>

                <!-- Title -->
                <TextBlock Grid.Row="0" Text="Load Calculation" FontSize="20" FontWeight="Bold" 
                           HorizontalAlignment="Center" Margin="0,0,0,10" />

                <!-- DataGrid -->
                <DataGrid Name="detailsGrid" AutoGenerateColumns="False" Height="300" Background="White" 
                          AllowDrop="True"
                          PreviewMouseLeftButtonDown="detailsGrid_mouse_down" 
                          DragEnter="detailsGrid_drag_enter" 
                          Drop="detailsGrid_drop">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="Element 1" Binding="{Binding Element1}" Width="*" />
                        <DataGridTextColumn Header="Element 2" Binding="{Binding Element2}" Width="*" />
                        <DataGridTextColumn Header="Element 3" Binding="{Binding Element3}" Width="*" />
                    </DataGrid.Columns>
                </DataGrid>

                <!-- Save Button -->
                <Button Grid.Row="2" Content="Save" Height="30" Width="80" 
                        HorizontalAlignment="Right" Margin="0,10,10,0" Click="save_data" />
            </Grid>
        </Window>
        """

        super(ParameterSetSelector, self).__init__(layout, literal_string=True)

        # Load data from the CSV file if it exists, otherwise use default data
        if os.path.exists(self.CSV_FILE):
            self.data = self.load_data_from_csv()
        else:
            # Initialize Calculation Data as dictionaries
            self.columns = {
                "Element1": ["Load Name", "Type", "Demand"],
                "Element2": ["Value", "Category", "Total Load"],
                "Element3": ["Voltage", "Current", "Power"]
            }
            self.data = self.convert_columns_to_rows()

        # Populate DataGrid
        self.converted_data = ObservableCollection[ExpandoObject]()
        self.refresh_grid()

        # Drag state
        self.dragged_item = None
        self.dragged_column = None

    def load_data_from_csv(self):
        """Load data from the CSV file."""
        data = []
        try:
            with open(self.CSV_FILE, "r") as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    data.append({
                        "Element1": row.get("Element1", ""),
                        "Element2": row.get("Element2", ""),
                        "Element3": row.get("Element3", "")
                    })
        except Exception as e:
            forms.alert("Error loading data from CSV: {}".format(e))
        return data

    def convert_columns_to_rows(self):
        """Convert column dictionary data into row-based data for the grid."""
        max_length = max(len(self.columns[key]) for key in self.columns)
        rows = []
        for i in range(max_length):
            row = {
                "Element1": self.columns["Element1"][i] if i < len(self.columns["Element1"]) else "",
                "Element2": self.columns["Element2"][i] if i < len(self.columns["Element2"]) else "",
                "Element3": self.columns["Element3"][i] if i < len(self.columns["Element3"]) else ""
            }
            rows.append(row)
        return rows

    def refresh_grid(self):
        # Remove rows where all columns are empty
        self.data = [row for row in self.data if row.get("Element1", "") or row.get("Element2", "") or row.get("Element3", "")]

        # Refresh the ObservableCollection
        self.converted_data.Clear()
        for row in self.data:
            expando = ExpandoObject()
            expando.Element1 = row.get("Element1", "")
            expando.Element2 = row.get("Element2", "")
            expando.Element3 = row.get("Element3", "")
            self.converted_data.Add(expando)
        self.detailsGrid.ItemsSource = self.converted_data

    def remove_duplicates(self):
        """Remove duplicate entries from each column."""
        seen_element1 = set()
        seen_element2 = set()
        seen_element3 = set()
        unique_data = []

        for row in self.data:
            element1 = row["Element1"]
            element2 = row["Element2"]
            element3 = row["Element3"]

            # Check and keep unique entries for each column
            if element1 not in seen_element1 or element1 == "":
                seen_element1.add(element1)
            else:
                row["Element1"] = ""

            if element2 not in seen_element2 or element2 == "":
                seen_element2.add(element2)
            else:
                row["Element2"] = ""

            if element3 not in seen_element3 or element3 == "":
                seen_element3.add(element3)
            else:
                row["Element3"] = ""

            unique_data.append(row)

        self.data = unique_data

    def save_data(self, sender, e):
        """Save the current data to a CSV file."""
        try:
            with open(self.CSV_FILE, "w") as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=["Element1", "Element2", "Element3"])
                writer.writeheader()
                for row in self.data:
                    writer.writerow(row)
            forms.alert("Data saved to {}".format(self.CSV_FILE))
        except IOError as io_err:
            forms.alert("Error saving data: {}".format(io_err))

    def detailsGrid_mouse_down(self, sender, e):
        try:
            # Capture the item being dragged and the column
            hit_test_result = sender.InputHitTest(e.GetPosition(sender))
            if hasattr(hit_test_result, "DataContext") and hit_test_result.DataContext is not None:
                self.dragged_item = hit_test_result.DataContext

                # Determine the dragged column by position
                clicked_point = e.GetPosition(sender)
                total_width = 0
                for column in sender.Columns:
                    total_width += column.ActualWidth
                    if clicked_point.X <= total_width:
                        if column.Header == "Element 1":
                            self.dragged_column = "Element1"
                        elif column.Header == "Element 2":
                            self.dragged_column = "Element2"
                        elif column.Header == "Element 3":
                            self.dragged_column = "Element3"
                        break

                # Validate the dragged data
                if self.dragged_column and getattr(self.dragged_item, self.dragged_column, None):
                    drag_data = DataObject()
                    drag_data.SetData("Text", getattr(self.dragged_item, self.dragged_column, ""))
                    DragDrop.DoDragDrop(sender, drag_data, DragDropEffects.Move)
                else:
                    forms.alert("Invalid column or data for drag operation.\nDragged Item: {}\nDragged Column: {}".format(self.dragged_item, self.dragged_column))
        except Exception as ex:
            forms.alert("Error in mouse down: " + str(ex))

    def detailsGrid_drag_enter(self, sender, e):
        try:
            if e.Data.GetDataPresent("Text"):
                e.Effects = DragDropEffects.Move
            else:
                e.Effects = DragDropEffects.None
            e.Handled = True
        except Exception as ex:
            forms.alert("Error in drag enter: " + str(ex))

    def detailsGrid_drop(self, sender, e):
        try:
            if self.dragged_item and self.dragged_column and e.Data.GetDataPresent("Text"):
                dropped_value = e.Data.GetData("Text")

                # Determine the target column
                target_column = None
                clicked_point = e.GetPosition(sender)
                total_width = 0
                for column in sender.Columns:
                    total_width += column.ActualWidth
                    if clicked_point.X <= total_width:
                        if column.Header == "Element 1":
                            target_column = "Element1"
                        elif column.Header == "Element 2":
                            target_column = "Element2"
                        elif column.Header == "Element 3":
                            target_column = "Element3"
                        break

                # Add a new row to the bottom of the target column with the dragged value
                if target_column:
                    new_row = {"Element1": "", "Element2": "", "Element3": ""}
                    new_row[target_column] = dropped_value
                    self.data.append(new_row)

                    # Update self.data directly to remove the dragged value from the original column
                    for row in self.data:
                        if row[self.dragged_column] == dropped_value:
                            row[self.dragged_column] = ""
                            break

                    # Remove duplicates from the columns
                    self.remove_duplicates()

                    # Refresh the grid
                    self.refresh_grid()

                # Clear drag state
                self.dragged_item = None
                self.dragged_column = None

                # Remove rows and cells that are completely blank
                self.data = [row for row in self.data if any(row.values())]
                self.refresh_grid()

            e.Handled = True
        except Exception as ex:
            forms.alert("Error in drop: " + str(ex))

try:
    selector = ParameterSetSelector()
    selector.show_dialog()
except Exception as e:
    forms.alert("Critical error: {}".format(e))
