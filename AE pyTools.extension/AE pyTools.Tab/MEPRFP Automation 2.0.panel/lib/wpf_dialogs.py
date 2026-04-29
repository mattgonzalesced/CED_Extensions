# -*- coding: utf-8 -*-
"""
Lightweight WPF-based prompts: single-line text input and list picker.

Separated from ``forms_compat`` so a module-load failure in the WPF
stack only affects scripts that genuinely need WPF (the Stage 1
authoring tools), not the Stage 3 Import/Export scripts that only
touch Windows.Forms.

API:
    prompt_for_string(prompt, title="", default="")  ->  str | None
    pick_from_list(options, title="Pick one", prompt="Choose:",
                   display_func=None)               ->  T   | None
"""

import clr  # noqa: F401

clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System.IO import StringReader  # noqa: E402
from System.Windows.Markup import XamlReader  # noqa: E402
from System.Xml import XmlReader  # noqa: E402


def _load_xaml(text):
    return XamlReader.Load(XmlReader.Create(StringReader(text)))


_STRING_PROMPT_XAML = """\
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="" Width="460" Height="170"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
  <Grid Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Grid.Row="0" x:Name="PromptText" TextWrapping="Wrap" Margin="0,0,0,8"/>
    <TextBox  Grid.Row="1" x:Name="InputBox" Margin="0,0,0,8"/>
    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
      <Button x:Name="OkButton" Content="OK" Width="80" Margin="0,0,8,0" IsDefault="True"/>
      <Button x:Name="CancelButton" Content="Cancel" Width="80" IsCancel="True"/>
    </StackPanel>
  </Grid>
</Window>
"""

_LIST_PICKER_XAML = """\
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="" Width="520" Height="460"
        WindowStartupLocation="CenterScreen">
  <Grid Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Grid.Row="0" x:Name="PromptText" Margin="0,0,0,8"/>
    <ListBox   Grid.Row="1" x:Name="Options"/>
    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
      <Button x:Name="OkButton" Content="OK" Width="80" Margin="0,0,8,0" IsDefault="True"/>
      <Button x:Name="CancelButton" Content="Cancel" Width="80" IsCancel="True"/>
    </StackPanel>
  </Grid>
</Window>
"""


_MULTI_SELECT_XAML = """\
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="" Width="640" Height="540"
        WindowStartupLocation="CenterScreen">
  <Window.Resources>
    <Style x:Key="CheckListBoxItem" TargetType="ListBoxItem">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ListBoxItem">
            <Border Background="Transparent" Padding="2">
              <CheckBox IsChecked="{Binding IsSelected,
                                            RelativeSource={RelativeSource TemplatedParent},
                                            Mode=TwoWay}"
                        VerticalContentAlignment="Center">
                <ContentPresenter VerticalAlignment="Center"/>
              </CheckBox>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>
  <Grid Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Grid.Row="0" x:Name="PromptText" Margin="0,0,0,8" TextWrapping="Wrap"/>
    <ListBox   Grid.Row="1" x:Name="Options"
               SelectionMode="Multiple"
               ItemContainerStyle="{StaticResource CheckListBoxItem}"/>
    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
      <Button x:Name="CheckAllButton" Content="Check all" Width="100" Margin="0,0,8,0"/>
      <Button x:Name="UncheckAllButton" Content="Uncheck all" Width="100" Margin="0,0,16,0"/>
      <Button x:Name="OkButton" Content="OK" Width="80" Margin="0,0,8,0" IsDefault="True"/>
      <Button x:Name="CancelButton" Content="Cancel" Width="80" IsCancel="True"/>
    </StackPanel>
  </Grid>
</Window>
"""


class _StringPromptDialog(object):
    def __init__(self, prompt, title, default):
        self.window = _load_xaml(_STRING_PROMPT_XAML)
        self.window.Title = title or ""
        self.window.FindName("PromptText").Text = prompt or ""
        self._input = self.window.FindName("InputBox")
        self._input.Text = default or ""
        self._input.SelectAll()
        self.window.FindName("OkButton").Click += self._on_ok
        self.window.FindName("CancelButton").Click += self._on_cancel
        self._result = None

    def _on_ok(self, sender, e):
        self._result = self._input.Text or ""
        self.window.Close()

    def _on_cancel(self, sender, e):
        self._result = None
        self.window.Close()

    def show(self):
        self.window.ShowDialog()
        return self._result


class _ListPickerDialog(object):
    def __init__(self, options, title, prompt, display_func):
        self.window = _load_xaml(_LIST_PICKER_XAML)
        self.window.Title = title or ""
        self.window.FindName("PromptText").Text = prompt or ""
        self._listbox = self.window.FindName("Options")
        self._options = list(options or [])
        for opt in self._options:
            label = display_func(opt) if display_func else str(opt)
            self._listbox.Items.Add(label)
        self.window.FindName("OkButton").Click += self._on_ok
        self.window.FindName("CancelButton").Click += self._on_cancel
        self._result = None

    def _on_ok(self, sender, e):
        idx = self._listbox.SelectedIndex
        if idx >= 0:
            self._result = self._options[idx]
        self.window.Close()

    def _on_cancel(self, sender, e):
        self._result = None
        self.window.Close()

    def show(self):
        self.window.ShowDialog()
        return self._result


def prompt_for_string(prompt, title="", default=""):
    return _StringPromptDialog(prompt, title, default).show()


def pick_from_list(options, title="Pick one", prompt="Choose:", display_func=None):
    return _ListPickerDialog(options, title, prompt, display_func).show()


def multi_select_from_list(options, title="Pick options", prompt="Check items:",
                           display_func=None):
    """Multi-select picker with checkboxes. Returns the list of chosen
    objects (in the order they were checked), or ``None`` if cancelled.
    Uses the same checkbox-templated ListBox pattern as the placement
    filters so click-anywhere-on-row toggles selection.
    """
    return _MultiSelectDialog(options, title, prompt, display_func).show()


class _MultiSelectDialog(object):
    def __init__(self, options, title, prompt, display_func):
        self.window = _load_xaml(_MULTI_SELECT_XAML)
        self.window.Title = title or ""
        self.window.FindName("PromptText").Text = prompt or ""
        self._listbox = self.window.FindName("Options")
        self._options = list(options or [])
        for opt in self._options:
            label = display_func(opt) if display_func else str(opt)
            self._listbox.Items.Add(label)
        # Retained handlers (pythonnet GC defence).
        self._h_ok = lambda s, e: self._on_ok(s, e)
        self._h_cancel = lambda s, e: self._on_cancel(s, e)
        self._h_check_all = lambda s, e: self._on_check_all(s, e)
        self._h_uncheck_all = lambda s, e: self._on_uncheck_all(s, e)
        self.window.FindName("OkButton").Click += self._h_ok
        self.window.FindName("CancelButton").Click += self._h_cancel
        self.window.FindName("CheckAllButton").Click += self._h_check_all
        self.window.FindName("UncheckAllButton").Click += self._h_uncheck_all
        self._result = None

    def _on_ok(self, sender, e):
        chosen = []
        # SelectedItems holds the *labels* (since we added strings).
        # Map back to options by label.
        selected_labels = list(self._listbox.SelectedItems)
        if selected_labels:
            label_to_option = {}
            for opt in self._options:
                label = str(opt) if not callable(getattr(opt, "to_label", None)) else opt.to_label()
                label_to_option.setdefault(label, opt)
            # Better: rebuild from the items list directly.
            for i in range(self._listbox.Items.Count):
                container = self._listbox.ItemContainerGenerator.ContainerFromIndex(i)
                if container is not None and getattr(container, "IsSelected", False):
                    chosen.append(self._options[i])
        self._result = chosen
        self.window.Close()

    def _on_cancel(self, sender, e):
        self._result = None
        self.window.Close()

    def _on_check_all(self, sender, e):
        self._listbox.SelectAll()

    def _on_uncheck_all(self, sender, e):
        self._listbox.UnselectAll()

    def show(self):
        self.window.ShowDialog()
        return self._result
