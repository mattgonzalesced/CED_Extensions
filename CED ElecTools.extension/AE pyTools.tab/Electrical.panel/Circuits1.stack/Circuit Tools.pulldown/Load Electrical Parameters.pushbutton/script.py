# -*- coding: utf-8 -*-

from pyrevit import forms, revit, script

from CEDElectrical.Domain import settings_manager

doc = revit.doc
logger = script.get_logger()
output = script.get_output()


def _format_summary(result):
    warnings = list((result or {}).get("warnings") or [])
    errors = list((result or {}).get("errors") or [])
    locked = list((result or {}).get("locked") or [])

    lines = [
        "Updated: {}".format(int((result or {}).get("updated") or 0)),
        "Unchanged: {}".format(int((result or {}).get("unchanged") or 0)),
        "Skipped: {}".format(int((result or {}).get("skipped") or 0)),
        "Category unbind updates: {}".format(int((result or {}).get("unbound") or 0)),
    ]

    if locked:
        lines.append("Skipped (owned by others): {}".format(int(len(locked))))
    if warnings:
        lines.append("Warnings: {}".format(int(len(warnings))))
    if errors:
        lines.append("Errors: {}".format(int(len(errors))))

    return "\n".join(lines), warnings, errors, locked


def main():
    if doc is None:
        forms.alert("No active Revit document found.")
        return

    output.close_others()
    output.show()
    output.print_md("## Load Electrical Parameters")

    settings = settings_manager.load_circuit_settings(doc)
    result = settings_manager.sync_electrical_parameter_bindings(
        doc,
        logger=logger,
        settings=settings,
        check_ownership=True,
        transaction_name="Load Electrical Parameters",
    )

    status = str((result or {}).get("status") or "").lower()
    summary_text, warnings, errors, locked = _format_summary(result)

    if status == "failed":
        forms.alert("Load Electrical Parameters failed.\n\n{}".format((result or {}).get("reason") or "Unknown error."))
        if errors:
            output.print_md("### Errors")
            for message in errors:
                output.print_md("- {}".format(message))
        return

    output.print_md("### Summary")
    for line in summary_text.splitlines():
        output.print_md("- {}".format(line))

    if warnings:
        output.print_md("### Warnings")
        for message in warnings:
            output.print_md("- {}".format(message))

    if locked:
        output.print_md("### Owned By Other User")
        for item in locked:
            output.print_md(
                "- {} (owner: {})".format(
                    str(item.get("parameter") or "Unnamed Parameter"),
                    str(item.get("owner") or "Unknown"),
                )
            )

    if errors:
        output.print_md("### Errors")
        for message in errors:
            output.print_md("- {}".format(message))
        forms.alert("Load completed with errors.\n\n{}".format(summary_text))
        return

    forms.alert("Load Electrical Parameters completed.\n\n{}".format(summary_text))


if __name__ == "__main__":
    main()
