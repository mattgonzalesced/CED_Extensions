# -*- coding: utf-8 -*-
"""
PlaceByRoomUI
-------------
WPF window to capture per-room profile and type counts.
"""

from pyrevit import forms
from System.Windows import Thickness, GridLength, GridUnitType, FontWeights
from System.Windows.Controls import (
    Grid,
    RowDefinition,
    ColumnDefinition,
    TextBlock,
    TextBox,
    StackPanel,
    Separator,
    Orientation,
)


class PlaceByRoomWindow(forms.WPFWindow):
    def __init__(self, xaml_path, rooms, profiles, type_map):
        forms.WPFWindow.__init__(self, xaml_path)
        self._rooms = rooms or []
        self._profiles = profiles or []
        self._type_map = type_map or {}
        self._profile_boxes = {}
        self._type_boxes = {}
        self.counts = None

        header = self.FindName("HeaderText")
        if header is not None:
            header.Text = "Set profile and type counts per room. Profile counts multiply type counts."

        self._build_panel()

        apply_btn = self.FindName("ApplyButton")
        cancel_btn = self.FindName("CancelButton")
        if apply_btn is not None:
            apply_btn.Click += self._on_apply
        if cancel_btn is not None:
            cancel_btn.Click += self._on_cancel

    def _build_panel(self):
        panel = self.FindName("RoomPanel")
        if panel is None:
            return

        for idx, room in enumerate(self._rooms):
            room_id = room.get("id")
            room_label = room.get("label") or "<Room>"

            if idx:
                panel.Children.Add(Separator())

            header = TextBlock()
            header.Text = "Room: {}".format(room_label)
            header.FontWeight = FontWeights.Bold
            header.Margin = Thickness(0, 8, 0, 4)
            panel.Children.Add(header)

            for profile in self._profiles:
                labels = self._type_map.get(profile) or []

                profile_panel = StackPanel()
                profile_panel.Margin = Thickness(8, 2, 0, 8)

                profile_header = TextBlock()
                profile_header.Text = "Profile: {}".format(profile)
                profile_header.FontWeight = FontWeights.SemiBold
                profile_header.Margin = Thickness(0, 0, 0, 2)
                profile_panel.Children.Add(profile_header)

                count_row = StackPanel()
                count_row.Orientation = Orientation.Horizontal
                count_row.Margin = Thickness(0, 0, 0, 4)
                label = TextBlock()
                label.Text = "Profile count:"
                label.Width = 120
                count_row.Children.Add(label)
                profile_box = TextBox()
                profile_box.Text = "1"
                profile_box.Width = 60
                count_row.Children.Add(profile_box)
                profile_panel.Children.Add(count_row)
                self._profile_boxes[(room_id, profile)] = profile_box

                if labels:
                    grid = Grid()
                    grid.Margin = Thickness(0, 0, 0, 2)

                    col_type = ColumnDefinition()
                    col_type.Width = GridLength(420, GridUnitType.Pixel)
                    grid.ColumnDefinitions.Add(col_type)
                    col_count = ColumnDefinition()
                    col_count.Width = GridLength(80, GridUnitType.Pixel)
                    grid.ColumnDefinitions.Add(col_count)

                    header_row = RowDefinition()
                    header_row.Height = GridLength(22, GridUnitType.Pixel)
                    grid.RowDefinitions.Add(header_row)

                    type_header = TextBlock()
                    type_header.Text = "Type"
                    type_header.FontWeight = FontWeights.Bold
                    type_header.Margin = Thickness(0, 0, 6, 2)
                    Grid.SetRow(type_header, 0)
                    Grid.SetColumn(type_header, 0)
                    grid.Children.Add(type_header)

                    count_header = TextBlock()
                    count_header.Text = "Count"
                    count_header.FontWeight = FontWeights.Bold
                    count_header.Margin = Thickness(0, 0, 6, 2)
                    Grid.SetRow(count_header, 0)
                    Grid.SetColumn(count_header, 1)
                    grid.Children.Add(count_header)

                    row_idx = 1
                    for label_text in labels:
                        row = RowDefinition()
                        row.Height = GridLength(26, GridUnitType.Pixel)
                        grid.RowDefinitions.Add(row)

                        type_label = TextBlock()
                        type_label.Text = label_text
                        type_label.Margin = Thickness(0, 0, 6, 2)
                        Grid.SetRow(type_label, row_idx)
                        Grid.SetColumn(type_label, 0)
                        grid.Children.Add(type_label)

                        count_box = TextBox()
                        count_box.Text = "0"
                        count_box.Width = 60
                        Grid.SetRow(count_box, row_idx)
                        Grid.SetColumn(count_box, 1)
                        grid.Children.Add(count_box)

                        self._type_boxes[(room_id, profile, label_text)] = count_box
                        row_idx += 1

                    profile_panel.Children.Add(grid)
                else:
                    note = TextBlock()
                    note.Text = "No linked types found for this profile."
                    note.Margin = Thickness(0, 0, 0, 2)
                    profile_panel.Children.Add(note)

                panel.Children.Add(profile_panel)

    def _parse_count(self, value):
        if value is None:
            return None
        raw = str(value).strip()
        if not raw:
            return None
        try:
            count = int(raw)
        except Exception:
            return None
        if count < 0:
            return None
        return count

    def _on_apply(self, sender, args):
        counts = {}
        for room in self._rooms:
            room_id = room.get("id")
            room_label = room.get("label") or "<Room>"
            counts.setdefault(room_id, {"profile": {}, "types": {}})
            for profile in self._profiles:
                profile_box = self._profile_boxes.get((room_id, profile))
                profile_count = self._parse_count(getattr(profile_box, "Text", None))
                if profile_count is None:
                    forms.alert(
                        "Enter a whole number zero or greater for profile '{}' in room '{}'."
                        .format(profile, room_label),
                        title="Place by Room Counts",
                    )
                    return
                counts[room_id]["profile"][profile] = profile_count
                for label in self._type_map.get(profile) or []:
                    type_box = self._type_boxes.get((room_id, profile, label))
                    type_count = self._parse_count(getattr(type_box, "Text", None))
                    if type_count is None:
                        forms.alert(
                            "Enter a whole number zero or greater for type '{}' in room '{}'."
                            .format(label, room_label),
                            title="Place by Room Counts",
                        )
                        return
                    counts[room_id]["types"].setdefault(profile, {})[label] = type_count
        self.counts = counts
        self.DialogResult = True
        self.Close()

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()
