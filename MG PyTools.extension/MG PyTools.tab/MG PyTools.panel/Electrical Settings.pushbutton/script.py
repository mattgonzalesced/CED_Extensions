# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.DB.Electrical import ElectricalSetting

# Get the active Revit document from PyRevit

def set_circuits():
    try:
        doc = __revit__.ActiveUIDocument.Document
        # Access electrical settings
        electrical_settings = ElectricalSetting.GetElectricalSettings(doc)

        # Get current value
        current_sequence = electrical_settings.CircuitSequence
        current_value = current_sequence.ToString()

        print("Current circuit sequence setting: {}".format(current_value))

        # Only update if not already OddThenEven
        if current_value != "OddThenEven":
            t = Transaction(doc, "Set Circuit Sequence to OddThenEven")
            t.Start()
            # Set the sequence (enum value via string comparison workaround)
            electrical_settings.CircuitSequence = current_sequence.__class__.OddThenEven
            t.Commit()
            print("Circuit sequence was updated to 'OddThenEven'.")
        else:
            print("No changes needed. Already set to 'OddThenEven'.")

    except Exception as e:
        print("Failed to read or update electrical settings:\n{}".format(str(e)))

set_circuits()