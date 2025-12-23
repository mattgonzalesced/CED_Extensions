# -*- coding: utf-8 -*-
from pyrevit import script, revit
from CEDElectrical.Domain.one_line_tree import build_system_tree

logger = script.get_logger()


def print_tree(output, tree):
    def visitor(item, level):
        indent = "    " * level
        if hasattr(item, "circuit_number"):
            output.print_md("{}- Circuit `{}` | Load: `{}`".format(
                indent, item.circuit_number, item.load_name
            ))
        else:
            output.print_md("{}**{}** (ID: {}) _({})_".format(
                indent, item.panel_name, item.element_id, item.equipment_type
            ))

    for root in tree.root_nodes:
        output.print_md("---\n### Root: **{}**".format(root.panel_name))
        tree.walk_tree(root, visitor)


if __name__ == "__main__":
    output = script.get_output()
    output.close_others()
    tree = build_system_tree(revit.doc)
    print_tree(output, tree)
