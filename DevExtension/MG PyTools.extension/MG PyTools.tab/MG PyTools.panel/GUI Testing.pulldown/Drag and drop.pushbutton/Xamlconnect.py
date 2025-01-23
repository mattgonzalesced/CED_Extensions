import clr
clr.AddReference('system.windows.forms')
clr.AddReference('IronPython.wpf')

import clr
clr.AddReference('PresentationFramework')  # Required for WPF functionality
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from pyrevit import script
from System.Windows import Window
import wpf

from pyrevit import script
xamlfile = script.get_bundle_file('ui.xaml')

import wpf
from system import Windows

class MyWindow(Windows.Window):
    def __init__(self):
        wpf.LoadComponent (self,xamlfile)


###lets show the window
MyWindow().ShowDialog()


