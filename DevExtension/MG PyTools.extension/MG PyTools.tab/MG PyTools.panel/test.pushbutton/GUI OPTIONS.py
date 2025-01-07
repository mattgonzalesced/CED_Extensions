from pyrevit import forms

# The vending machine layout with a color picker
layout = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Snack Selector with Color Picker"
        Height="400" Width="400"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize">
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
        <TextBlock Grid.Row="0" Text="Choose Your Snacks and a Color" FontSize="20" FontWeight="Bold" Foreground="White" 
                   HorizontalAlignment="Center" Margin="0,0,0,10" />

        <!-- Checkboxes -->
        <StackPanel Grid.Row="1" VerticalAlignment="Top" Margin="20">
            <CheckBox Name="checkbox1" Content="Chips" FontSize="16" Foreground="White" Margin="0,10" />
            <CheckBox Name="checkbox2" Content="Chocolate" FontSize="16" Foreground="White" Margin="0,10" />
            <CheckBox Name="checkbox3" Content="Cookies" FontSize="16" Foreground="White" Margin="0,10" />
            <TextBlock Text="Pick a color:" FontSize="14" Foreground="White" Margin="0,20,0,5" />
            <ComboBox Name="colorPicker" SelectedIndex="0" Margin="0,5">
                <ComboBoxItem Content="Gray (#FF808080)" />
                <ComboBoxItem Content="Red (#FFFF0000)" />
                <ComboBoxItem Content="Blue (#FF0000FF)" />
                <ComboBoxItem Content="Green (#FF008000)" />
            </ComboBox>
        </StackPanel>

        <!-- Button -->
        <Button Grid.Row="2" Name="okButton" Content="Confirm" FontSize="16" FontWeight="Bold"
                HorizontalAlignment="Center" Width="120" Background="#FF64B5F6" Foreground="White"
                Margin="0,10" />
    </Grid>
</Window>
"""

# Custom vending machine class
class SnackSelector(forms.WPFWindow):
    def __init__(self):
        super(SnackSelector, self).__init__(layout, literal_string=True)
        self.okButton.Click += self.on_done_clicked  # Connect "Done" button to action
        self.selected_snacks = []  # To store selected snacks
        self.selected_color = None  # To store the selected color

    def on_done_clicked(self, sender, args):
        # Check which snacks are selected
        if self.checkbox1.IsChecked:
            self.selected_snacks.append("Chips")
        if self.checkbox2.IsChecked:
            self.selected_snacks.append("Chocolate")
        if self.checkbox3.IsChecked:
            self.selected_snacks.append("Cookies")
        
        # Get the selected color
        selected_color_item = self.colorPicker.SelectedItem
        if selected_color_item:
            self.selected_color = selected_color_item.Content.split(' ')[-1]  # Extract the hex color code

        self.Close()  # Close the vending machine (window)

# Using the vending machine
machine = SnackSelector()  # Create the vending machine
machine.show_dialog()  # Open the vending machine and wait until the user closes it

# Handle results AFTER the user closes the window
if machine.selected_snacks or machine.selected_color:
    snacks = ", ".join(machine.selected_snacks) if machine.selected_snacks else "No snacks"
    color = machine.selected_color if machine.selected_color else "No color"
    forms.alert("You selected:\nSnacks: {0}\nColor: {1}".format(snacks, color), title="Your Selection")
else:
    forms.alert("No selection made.", title="No Selection")
