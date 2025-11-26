# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms, script
from pyrevit.framework import wpf
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import List

logger = script.get_logger()


# -----------------------------------------------------------------------------
# Data containers
# -----------------------------------------------------------------------------
class TreeBranch(object):
    def __init__(self, system, parent_node):
        self.system = system
        self.parent = parent_node
        self.circuit_number = system.CircuitNumber or ""
        self.load_name = system.LoadName or ""
        self.base_name = DB.Element.Name.__get__(system.BaseEquipment) if system.BaseEquipment else ""
        self.child_nodes = []
        self.connected_elements = []

    @property
    def search_text(self):
        child_names = ", ".join([c.panel_name for c in self.child_nodes])
        device_names = ", ".join([safe_element_name(el) for el in self.connected_elements])
        return " ".join([self.circuit_number, self.load_name, self.base_name, child_names, device_names]).lower()

    def build_label(self, level):
        target = ", ".join([c.panel_name for c in self.child_nodes]) if self.child_nodes else (self.load_name or "Loads")
        indent = "    " * level
        return u"{}{} → {}".format(indent, self.circuit_number or "(no number)", target)

    def build_tooltip(self):
        parts = [
            "Circuit: {}".format(self.circuit_number or "(not set)"),
            "From: {}".format(self.base_name or "Unknown"),
            "Load: {}".format(self.load_name or ""),
        ]
        if self.child_nodes:
            parts.append("Feeds: {}".format(", ".join([c.panel_name for c in self.child_nodes])))
        if self.connected_elements:
            parts.append("Devices: {}".format(", ".join([safe_element_name(el) for el in self.connected_elements])))
        return "\n".join(parts)


def safe_element_name(element):
    try:
        return DB.Element.Name.__get__(element)
    except Exception:
        return str(element.Id) if hasattr(element, "Id") else "Unknown"


class TreeNode(object):
    def __init__(self, element):
        self.element = element
        self.element_id = element.Id
        self.panel_name = safe_element_name(element)
        self.upstream = []
        self.downstream = []

    def collect_branches(self, nodes_map):
        mep = getattr(self.element, "MEPModel", None)
        if not mep:
            return

        systems = mep.GetElectricalSystems()
        if not systems:
            return

        for sys in systems:
            if sys.BaseEquipment and sys.BaseEquipment.Id == self.element_id:
                branch = TreeBranch(sys, self)
                elements = list(sys.Elements) if hasattr(sys, "Elements") and sys.Elements else []
                for el in elements:
                    if el.Category and int(el.Category.Id.IntegerValue) == int(DB.BuiltInCategory.OST_ElectricalEquipment):
                        child = nodes_map.get(el.Id)
                        if child and child not in branch.child_nodes:
                            branch.child_nodes.append(child)
                            child.upstream.append(self)
                    else:
                        branch.connected_elements.append(el)
                self.downstream.append(branch)


# -----------------------------------------------------------------------------
# Tree building helpers
# -----------------------------------------------------------------------------
def get_equipment_nodes(doc):
    nodes = {}
    collector = DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment) \
        .WhereElementIsNotElementType()
    for equip in collector:
        nodes[equip.Id] = TreeNode(equip)
    return nodes


def build_tree(doc):
    nodes = get_equipment_nodes(doc)
    for node in nodes.values():
        node.collect_branches(nodes)

    roots = []
    for node in nodes.values():
        if not node.upstream:
            roots.append(node)
    if not roots:
        roots = list(nodes.values())
    return roots


def flatten_tree(nodes):
    items = []
    visited = set()

    def _walk(node, level=0):
        if node.element_id in visited:
            return
        visited.add(node.element_id)
        for branch in node.downstream:
            items.append({
                "branch": branch,
                "level": level,
            })
            for child in branch.child_nodes:
                _walk(child, level + 1)

    for root in nodes:
        _walk(root, 0)
    return items


class TreeListItem(object):
    def __init__(self, branch, level):
        self.branch = branch
        self.level = level
        self.Label = branch.build_label(level)
        self.Tooltip = branch.build_tooltip()
        self.SearchKey = (self.Label + " " + branch.search_text).lower()

    @property
    def circuit(self):
        return self.branch.system

    @property
    def devices(self):
        return self.branch.connected_elements or []


# -----------------------------------------------------------------------------
# UI
# -----------------------------------------------------------------------------
DockableBase = getattr(forms, "ReactiveDockablePane", forms.WPFWindow)


class PowerSystemPane(DockableBase):
    PANE_ID = "ced.power.system.tree"

    def __init__(self):
        xamlfile = script.get_bundle_file("PowerSystemPane.xaml")
        if DockableBase is forms.WPFWindow:
            DockableBase.__init__(self, xamlfile)
        else:
            DockableBase.__init__(self, self.PANE_ID)
            wpf.LoadComponent(self, xamlfile)
            try:
                self.set_title("Power System Tree")
            except Exception:
                pass

        self.doc = revit.doc
        self.uidoc = revit.uidoc

        self.SearchBox = self.FindName("SearchBox")
        self.TreeList = self.FindName("TreeList")
        self.SelectCircuitsButton = self.FindName("SelectCircuitsButton")
        self.SelectDevicesButton = self.FindName("SelectDevicesButton")
        self.SelectBothButton = self.FindName("SelectBothButton")

        self.all_items = []
        # IronPython struggles to construct a typed ObservableCollection; use a non-generic
        # collection so it can host TreeListItem instances without throwing
        # "SystemError: No callable method" during initialization.
        self.visible_items = ObservableCollection()
        self.TreeList.ItemsSource = self.visible_items

        self._wire_events()
        self.refresh_items()

    def _wire_events(self):
        if self.SearchBox:
            self.SearchBox.TextChanged += self._on_search
        if self.SelectCircuitsButton:
            self.SelectCircuitsButton.Click += self._on_select_circuits
        if self.SelectDevicesButton:
            self.SelectDevicesButton.Click += self._on_select_devices
        if self.SelectBothButton:
            self.SelectBothButton.Click += self._on_select_both
        if self.TreeList:
            self.TreeList.MouseDoubleClick += self._on_double_click

    def refresh_items(self):
        try:
            roots = build_tree(self.doc)
            branches = flatten_tree(roots)
            self.all_items = [TreeListItem(data["branch"], data["level"]) for data in branches]
            self.apply_filter()
        except Exception as exc:
            logger.error("Failed to build power system tree: {}".format(exc))
            forms.alert("Unable to build the power system tree. Check the log for details.")

    def apply_filter(self):
        text = self.SearchBox.Text if self.SearchBox else ""
        query = text.lower() if text else ""
        self.visible_items.Clear()
        for item in self.all_items:
            if not query or query in item.SearchKey:
                self.visible_items.Add(item)

    def _on_search(self, sender, args):
        self.apply_filter()

    def _on_select_circuits(self, sender, args):
        self._select_items(include_circuits=True, include_devices=False)

    def _on_select_devices(self, sender, args):
        self._select_items(include_circuits=False, include_devices=True)

    def _on_select_both(self, sender, args):
        self._select_items(include_circuits=True, include_devices=True)

    def _on_double_click(self, sender, args):
        self._select_items(include_circuits=True, include_devices=True)

    def _select_items(self, include_circuits=False, include_devices=False):
        if not self.TreeList:
            return
        selected = list(self.TreeList.SelectedItems) if self.TreeList.SelectedItems else []
        if not selected:
            forms.alert("Select at least one circuit from the list.")
            return

        element_ids = set()
        if include_circuits:
            for item in selected:
                circuit = item.circuit
                if circuit:
                    element_ids.add(circuit.Id)
        if include_devices:
            for item in selected:
                for device in item.devices:
                    if device and hasattr(device, "Id"):
                        element_ids.add(device.Id)

        if not element_ids:
            forms.alert("No valid elements were found for the current selection.")
            return

        self.uidoc.Selection.SetElementIds(List[DB.ElementId](list(element_ids)))


# -----------------------------------------------------------------------------
# Entry point
# -----------------------------------------------------------------------------
def show_pane():
    pane = PowerSystemPane()
    try:
        pane.show()
    except Exception:
        try:
            pane.Show()
        except Exception:
            logger.warning("Unable to show dockable pane; falling back to modeless dialog.")
            try:
                pane.ShowDialog()
            except Exception:
                forms.alert("Unable to open the power system tree pane. Check the log for details.")


if __name__ == "__main__":
    show_pane()
