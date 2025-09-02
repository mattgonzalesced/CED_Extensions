# -*- coding: utf-8 -*-
# IronPython 2.7 / pyRevit / Revit API
# Sections reset lists. Levels:
#   A.       -> UpperCaseLetters, indent 1
#   1.       -> ArabicNumbers,   indent 2
#   1.A.     -> LowerCaseLetters, indent 3
# Removes true blank lines; preserves section headers; and appends '\v'
# to the end of any list paragraph that is followed by another list paragraph
# (visual gap without breaking numbering).

import re

from Autodesk.Revit.DB import Transaction, TextRange, ListType
from pyrevit import revit, DB, forms, script

LOGGER = script.get_logger()

# ---------- Patterns ----------
RE_SECTION_HEADER  = re.compile(r'^\s*\d{3,}\.\d+(\s|$)')         # e.g., 15100.02 INTENT
RE_NUM_DOT_LETTER  = re.compile(r'^\s*\d+\.\s*([A-Za-z])\.\s*')   # 1.A. or 1.a.
RE_NUMBER_ONLY     = re.compile(r'^\s*(\d+)\.\s*(?!\d)')          # 1.
RE_LETTER_ONLY     = re.compile(r'^\s*([A-Za-z])\.\s*')           # A. or a.

KIND_NONE = None
KIND_UC   = "UC"   # UpperCaseLetters (level 1)
KIND_AR   = "AR"   # ArabicNumbers   (level 2)
KIND_LC   = "LC"   # LowerCaseLetters (level 3)


def pick_single_textnote():
    uidoc = revit.uidoc
    doc   = revit.doc
    selected_ids = list(uidoc.Selection.GetElementIds())
    notes = []
    for elid in selected_ids:
        el = doc.GetElement(elid)
        if isinstance(el, DB.TextNote):
            notes.append(el)
    if len(notes) == 1:
        return notes[0]
    el = revit.pick_element(DB.BuiltInCategory.OST_TextNotes, "Select a Text Note to format")
    if el and isinstance(el, DB.TextNote):
        return el
    return None


def normalize_newlines(text_value):
    if not text_value:
        return u""
    if u"\r\n" in text_value:
        return text_value.replace(u"\r\n", u"\n")
    if u"\r" in text_value:
        return text_value.replace(u"\r", u"\n")
    return text_value


def denormalize_to_revit(text_with_lf):
    # Keep '\v' untouched; convert '\n' to '\r'
    return text_with_lf.replace(u"\n", u"\r")


def split_into_sections(lines):
    """Return list of { 'header': str, 'body': [str,...] } split at section headers."""
    sections = []
    current_header = None
    current_body = []
    saw_header = False

    def _push():
        if current_header is None and not current_body:
            return
        sections.append({"header": current_header, "body": list(current_body)})

    for line in lines:
        if RE_SECTION_HEADER.match(line):
            _push()
            current_header = line
            current_body = []
            saw_header = True
        else:
            if current_header is None and not saw_header:
                current_header = u""
            current_body.append(line)
    _push()
    return sections


def classify_and_strip(line_text):
    """Return (kind, indent, stripped_text) for one line."""
    m = RE_NUM_DOT_LETTER.match(line_text)
    if m:
        return (KIND_LC, 3, line_text[m.end():])
    m = RE_NUMBER_ONLY.match(line_text)
    if m:
        return (KIND_AR, 2, line_text[m.end():])
    m = RE_LETTER_ONLY.match(line_text)
    if m:
        return (KIND_UC, 1, line_text[m.end():])
    return (KIND_NONE, None, line_text)


def rebuild_text_and_meta(sections):
    """
    Build final paragraphs and per-paragraph metadata.
    - Keep section headers as plain paragraphs (None) which reset lists.
    - Remove true blank lines (prevents numbered empty items).
    Returns (final_lines, meta_per_line).
    """
    final_lines = []
    meta_per_line = []

    for section in sections:
        header = section["header"]
        body_lines = section["body"]

        if header is not None:
            header_text = header.strip()
            if header_text != u"":
                final_lines.append(header_text)
                meta_per_line.append({"kind": KIND_NONE, "indent": None})

        for raw in body_lines:
            if raw.strip() == u"":
                continue  # drop blanks entirely

            kind, indent, content = classify_and_strip(raw)

            if content.strip() == u"" and kind is not None:
                continue  # avoid empty list items

            final_lines.append(content)
            meta_per_line.append({"kind": kind, "indent": indent})

    return final_lines, meta_per_line


def add_vertical_tabs_at_list_runs(final_lines, meta_per_line):
    """
    Append '\v' inside any list paragraph that is immediately followed by another list paragraph.
    We do NOT add separate paragraphs; we append the char to the line content so numbering continues.
    Returns a new list of lines (same count) with some lines ending in '\v'.
    """
    augmented = []
    total = len(final_lines)
    for idx in range(total):
        line = final_lines[idx]
        kind = meta_per_line[idx]["kind"]
        if idx < total - 1:
            next_is_list = (meta_per_line[idx + 1]["kind"] is not None)
        else:
            next_is_list = False

        if (kind is not None) and next_is_list:
            # Avoid double-adding if it's already there
            if not line.endswith(u"\v"):
                line = line + u"\v"
        augmented.append(line)
    return augmented


def map_line_starts(lines_for_write):
    """Return the starting character index of each paragraph after joining with '\r'."""
    starts = []
    running = 0
    for text_line in lines_for_write:
        starts.append(running)
        if text_line == u"\v":
            # Not expected here (we append \v to lines, not standalone), but be safe.
            running += 1
        else:
            running += len(text_line) + 1  # include the trailing '\r'
    return starts


def apply_formatting_to_textnote(textnote, lines_for_write, meta_per_line):
    """Write content and apply list formatting per paragraph against the augmented lines."""
    text_with_lf = u"\n".join(lines_for_write)
    text_with_cr = denormalize_to_revit(text_with_lf)
    textnote.Text = text_with_cr

    starts = map_line_starts(lines_for_write)
    fmt = textnote.GetFormattedText()

    for i in range(len(lines_for_write)):
        kind = meta_per_line[i]["kind"]
        if kind is None:
            continue

        start_char = starts[i]
        # length = paragraph content length INCLUDING any appended '\v', plus the trailing '\r'
        length = len(lines_for_write[i]) + 1
        tr = TextRange(start_char, length)

        try:
            if kind == KIND_UC:
                fmt.SetListType(tr, ListType.UpperCaseLetters)
                fmt.SetIndentLevel(tr, 1)
            elif kind == KIND_AR:
                fmt.SetListType(tr, ListType.ArabicNumbers)
                fmt.SetIndentLevel(tr, 2)
            elif kind == KIND_LC:
                fmt.SetListType(tr, ListType.LowerCaseLetters)
                fmt.SetIndentLevel(tr, 3)
        except Exception as ex:
            LOGGER.warning("List format failed at paragraph {}: {}".format(i, ex))

    textnote.SetFormattedText(fmt)


def main():
    textnote = pick_single_textnote()
    if not textnote:
        forms.alert("Please select a single Text Note.", exitscript=True)
        return

    source_text = normalize_newlines(textnote.Text or u"")
    source_lines = source_text.split(u"\n")

    # 1) Split by section headers (each section resets lists)
    sections = split_into_sections(source_lines)

    # 2) Build paragraphs and metadata (indents and list kinds); no \v yet
    final_lines, meta_per_line = rebuild_text_and_meta(sections)

    # 3) Append '\v' at the end of any list line that is followed by another list line
    augmented_lines = add_vertical_tabs_at_list_runs(final_lines, meta_per_line)

    # 4) Write and format against the augmented lines (ranges account for the added '\v')
    with Transaction(revit.doc, "Format TextNote Lists (+\\v)") as tx:
        tx.Start()
        apply_formatting_to_textnote(textnote, augmented_lines, meta_per_line)
        tx.Commit()

    forms.alert("Done. Lists formatted by section; indicators removed; vertical tabs inserted.", ok=True)


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        LOGGER.exception("Formatting failed: {}".format(ex))
        forms.alert("Formatting failed:\n{}".format(ex), ok=True)
