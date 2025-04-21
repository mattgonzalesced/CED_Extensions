import re
from pyrevit import revit, DB, forms

doc = revit.doc
uidoc = revit.uidoc

# Get selected element IDs and filter for text notes
selected_ids = uidoc.Selection.GetElementIds()
text_notes = [doc.GetElement(el_id) for el_id in selected_ids if isinstance(doc.GetElement(el_id), DB.TextNote)]

if not text_notes:
    forms.alert("No text notes selected! Please select text notes and try again.", exitscript=True)

# Use the active view's bounding box to determine the top position of each text note.
def get_top_y(tn):
    bbox = tn.get_BoundingBox(doc.ActiveView)
    return bbox.Max.Y if bbox else 0

# Sort text notes by the top Y value (descending: top-most first)
sorted_text_notes = sorted(text_notes, key=get_top_y, reverse=True)

# Combine text from the sorted text notes using newline as separator.
combined_text = "\n".join([tn.Text for tn in sorted_text_notes])

# Replace any occurrence of 4 or more consecutive spaces with a single space.
combined_text = re.sub(r' {4,}', ' ', combined_text)

# Process newlines:
# A new line is started only if the line begins with:
# - A letter (A-Z or a-z) immediately followed by a period, OR
# - One or more digits followed by a period and a space.
pattern = r'^([A-Za-z]\.|[0-9]+\.\s)'

lines = combined_text.splitlines()
final_lines = []
for i, line in enumerate(lines):
    line = line.strip()
    if i == 0:
        final_lines.append(line)
    else:
        # If the line starts with the required pattern, it begins a new line.
        if re.match(pattern, line):
            final_lines.append(line)
        else:
            # Otherwise, append the line to the previous line.
            final_lines[-1] += " " + line
final_text = "\n".join(final_lines)

# Ask the user for a placement point for the merged text note.
placement_point = uidoc.Selection.PickPoint("Select a point to place the merged text note")

# Start a transaction group.
tgroup = DB.TransactionGroup(doc, "Merge Text Notes")
tgroup.Start()

try:
    # Transaction 1: Create the new merged text note.
    t1 = DB.Transaction(doc, "Create Merged Text Note")
    t1.Start()
    new_text_note = DB.TextNote.Create(
        doc,
        doc.ActiveView.Id,
        placement_point,
        final_text,
        sorted_text_notes[0].TextNoteType.Id
    )
    t1.Commit()

    # Transaction 2: Regenerate the document.
    t2 = DB.Transaction(doc, "Regenerate Document")
    t2.Start()
    doc.Regenerate()
    t2.Commit()


    # Transaction 4: Final regeneration.
    t4 = DB.Transaction(doc, "Final Regeneration")
    t4.Start()
    doc.Regenerate()
    t4.Commit()

    tgroup.Assimilate()

    forms.alert("Merged text notes into one successfully!")
    
except Exception as e:
    tgroup.RollBack()
    forms.alert("An error occurred: {}".format(e), exitscript=True)
