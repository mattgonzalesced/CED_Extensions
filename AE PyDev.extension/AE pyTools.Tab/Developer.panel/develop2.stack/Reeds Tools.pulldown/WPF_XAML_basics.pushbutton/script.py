# pyRevit WPF example: add two integers from UI inputs.
import clr
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xaml")

from pyrevit import script
xamlfile = script.get_bundle_file('ui.xaml')

# If 'import wpf' works in your environment, use it:
import wpf

from System.Windows import Window
from System.Windows.Media import Brushes

class MyWindow(Window):
    def __init__(self):
        # Inflate the Window from XAML
        wpf.LoadComponent(self, xamlfile)

        # Grab named elements
        self.AInput    = self.FindName('AInput')
        self.BInput    = self.FindName('BInput')
        self.AddButton = self.FindName('AddButton')
        self.StatusText = self.FindName('StatusText')
        self.AValueText = self.FindName('AValueText')
        self.BValueText = self.FindName('BValueText')
        self.SumValueText = self.FindName('SumValueText')

        # Wire events
        self.AddButton.Click += self.on_add_click

    def on_add_click(self, sender, e):
        # Clear status and reset borders
        self.StatusText.Text = ""
        for tb in (self.AInput, self.BInput):
            tb.ClearValue(type(tb).BorderBrushProperty)
            tb.ClearValue(type(tb).BorderThicknessProperty)

        a_txt = self.AInput.Text.strip()
        b_txt = self.BInput.Text.strip()

        # Parse as integers with simple validation
        try:
            a = int(a_txt)
        except Exception:
            self.StatusText.Text = "A must be an integer."
            self.AInput.BorderBrush = Brushes.Red
            self.AInput.BorderThickness = system_double(2)
            return

        try:
            b = int(b_txt)
        except Exception:
            self.StatusText.Text = "B must be an integer."
            self.BInput.BorderBrush = Brushes.Red
            self.BInput.BorderThickness = system_double(2)
            return

        s = a + b

        # Display results
        self.AValueText.Text = str(a)
        self.BValueText.Text = str(b)
        self.SumValueText.Text = str(s)

def system_double(val):
    # Helper for setting BorderThickness from IronPython
    from System import Double
    from System.Windows import Thickness
    return Thickness(Double(val))

# Show the dialog
MyWindow().ShowDialog()