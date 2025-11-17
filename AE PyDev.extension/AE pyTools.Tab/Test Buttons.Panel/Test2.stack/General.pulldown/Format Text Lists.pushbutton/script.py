# -*- coding: utf-8 -*-
# IronPython 2.7 / pyRevit / Revit API
# Sections reset lists. Levels:
#   A.       -> UpperCaseLetters, indent 1
#   1.       -> ArabicNumbers,   indent 2
#   1.A.     -> LowerCaseLetters, indent 3
# Removes true blank lines; preserves section headers;
# Optional vertical tabs can add visual spacing between list items.

import re

from Autodesk.Revit.DB import Transaction, TextRange, ListType
from pyrevit import revit, DB, forms, script

LOGGER = script.get_logger()


class PatternType(object):
    """Logical pattern types. You never touch regex directly."""
    UPPER = "UPPER"      # A. B. C. AA. AB.
    LOWER = "LOWER"      # a. b. c. (available if you want it)
    NUMERIC = "NUMERIC"  # 1. 2. 3.
    NUMDOT = "NUMDOT"    # 1.A. 1.a. 2.AA.


class ListLevel(object):
    def __init__(self, pattern_type, list_type, indent,
                 blank_before=False, blank_between=False, blank_after=False):
        """
        pattern_type : PatternType constant (UPPER, LOWER, NUMERIC, NUMDOT)
        list_type    : Revit ListType (UpperCaseLetters, ArabicNumbers, etc.)
        indent       : integer indent level for SetIndentLevel
        blank_* flags:
            blank_before  : real blank line before a *run* of this list level (not used yet)
            blank_between : use vertical tabs between items of this level
            blank_after   : real blank line after a run of this level (not used yet)
        """
        self.pattern_type = pattern_type
        self.list_type = list_type
        self.indent = indent
        self.blank_before = blank_before
        self.blank_between = blank_between
        self.blank_after = blank_after

        # Bind the appropriate regex based on pattern_type
        if pattern_type == PatternType.UPPER:
            # A. B. C. AA. AB. ...
            self.regex = re.compile(r'^\s*([A-Z]{1,3})\.\s*')
        elif pattern_type == PatternType.LOWER:
            # a. b. c. ...
            self.regex = re.compile(r'^\s*([a-z]{1,3})\.\s*')
        elif pattern_type == PatternType.NUMERIC:
            # 1. 2. 3. ...
            self.regex = re.compile(r'^\s*(\d+)\.\s*(?!\d)')
        elif pattern_type == PatternType.NUMDOT:
            # 1.A. 1.a. 2.AA. ...
            self.regex = re.compile(r'^\s*\d+\.\s*([A-Za-z]{1,3})\.\s*')
        else:
            # Fallback: never match
            self.regex = re.compile(r'^(?!)')

    def match(self, text):
        return self.regex.match(text)


# -------- SECTION HEADER PATTERN (for splitting & spacing) --------
# e.g. "15100.02 INTENT" or "3.01  INSTALLATION" or "1.07 SUBMITTALS"
RE_SECTION_HEADER = re.compile(r'^\s*\d{1,5}\.\d+(\s|$)')


# -------- LIST LEVEL CONFIG (this is the part you tweak) --------
LIST_LEVELS = [
    # Level 1: A. B. C. AA. AB. ...
    ListLevel(
        PatternType.UPPER,
        list_type=ListType.UpperCaseLetters,
        indent=1,
        blank_before=False,
        blank_between=False,   # NO vertical tabs between A. items
        blank_after=False
    ),

    # Level 2: 1. 2. 3. ...
    ListLevel(
        PatternType.NUMERIC,
        list_type=ListType.ArabicNumbers,
        indent=2,
        blank_before=False,
        blank_between=False,   # NO vertical tabs between 1. items
        blank_after=False
    ),

    # Level 3: 1.A. 1.a. 2.AA. ...
    ListLevel(
        PatternType.LOWER,
        list_type=ListType.LowerCaseLetters,
        indent=3,
        blank_before=False,
        blank_between=False,   # NO vertical tabs between 1.A. items
        blank_after=False
    ),
]


def pick_single_textnote():
    uidoc = revit.uidoc
    doc = revit.doc
    selected_ids = list(uidoc.Selection.GetElementIds())
    notes = []
    for element_id in selected_ids:
        element = doc.GetElement(element_id)
        if isinstance(element, DB.TextNote):
            notes.append(element)
    if len(notes) == 1:
        return notes[0]
    element = revit.pick_element(DB.BuiltInCategory.OST_TextNotes, "Select a Text Note to format")
    if element and isinstance(element, DB.TextNote):
        return element
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


def clean_internal_tabs(line_text):
    """
    Clean tabs INSIDE a list line (not headers).
    - Remove any tab characters.
    - Collapse runs of spaces/tabs to a single space.
    - Remove internal '\v' but preserve a trailing '\v' if present.
    """
    if not line_text:
        return line_text

    has_trailing_vt = line_text.endswith(u"\v")
    if has_trailing_vt:
        core = line_text[:-1]
    else:
        core = line_text

    # Replace any run of tabs/spaces with a single space
    core = re.sub(u'[ \t]+', u' ', core)
    # Remove internal vertical tabs (should not exist here)
    core = core.replace(u"\v", u"")

    if has_trailing_vt:
        return core + u"\v"
    return core


def split_into_sections(lines):
    """Return list of { 'header': str or None, 'body': [str,...] } split at numeric section headers."""
    sections = []
    current_header = None
    current_body = []
    saw_header = False

    def push_section():
        if current_header is None and not current_body:
            return
        sections.append({"header": current_header, "body": list(current_body)})

    for line in lines:
        if RE_SECTION_HEADER.match(line):
            push_section()
            current_header = line
            current_body = []
            saw_header = True
        else:
            if current_header is None and not saw_header:
                current_header = u""
            current_body.append(line)

    push_section()
    return sections


def classify_and_strip(line_text):
    """
    Try each configured ListLevel.
    Returns (level_object_or_None, indent_or_None, stripped_text).
    Tab cleaning ONLY happens for list content.
    """
    for level in LIST_LEVELS:
        match_obj = level.match(line_text)
        if match_obj:
            stripped = line_text[match_obj.end():]
            stripped = clean_internal_tabs(stripped)
            return level, level.indent, stripped
    return None, None, line_text


def rebuild_text_and_meta(sections):
    """
    Build final paragraphs and per-paragraph metadata.
    - Section headers become plain paragraphs (level None) and reset lists.
    - True blank lines from source are removed (prevents numbered empties).
    - Insert a real blank line before a numeric section header if the previous
      paragraph was a list paragraph.
    Returns (final_lines, meta_per_line) where meta_per_line[i]["level"] is
    either a ListLevel or None.
    """
    final_lines = []
    meta_per_line = []
    last_level = None  # level of the last appended paragraph (None for headers/plain text)

    for section in sections:
        header = section["header"]
        body_lines = section["body"]

        # ---- SECTION HEADER ----
        if header is not None:
            header_text = header.strip()
            if header_text != u"":
                # If last appended paragraph was a list, insert a blank line
                if last_level is not None and len(final_lines) > 0:
                    final_lines.append(u"")
                    meta_per_line.append({"level": None})
                    last_level = None

                final_lines.append(header_text)
                meta_per_line.append({"level": None})
                last_level = None

        # ---- SECTION BODY ----
        for raw_line in body_lines:
            if raw_line.strip() == u"":
                # drop blanks entirely from the source
                continue

            level, indent, content = classify_and_strip(raw_line)

            if content.strip() == u"" and level is not None:
                # avoid empty list items
                continue

            final_lines.append(content)
            meta_per_line.append({"level": level})
            last_level = level

    return final_lines, meta_per_line


def add_vertical_tabs_at_list_runs(final_lines, meta_per_line):
    """
    Optionally append '\v' to list paragraphs according to level.blank_between.
    - If blank_between is True and the next paragraph is also a list, we append '\v'.
    - Tabs are cleaned from list lines either way.
    Returns a new list of lines (same length as final_lines).
    """
    augmented = []
    total = len(final_lines)

    for index in range(total):
        line = final_lines[index]
        level = meta_per_line[index]["level"]

        if level is not None:
            # determine whether next is a list paragraph
            if index < total - 1:
                next_level = meta_per_line[index + 1]["level"]
            else:
                next_level = None

            # Always clean tabs for list lines
            line = clean_internal_tabs(line)

            # Only add '\v' if this level's config says so AND the next is also list
            if level.blank_between and next_level is not None:
                if not line.endswith(u"\v"):
                    line = line + u"\v"

        augmented.append(line)

    return augmented


def map_line_starts(lines_for_write):
    """Return the starting character index of each paragraph after joining with '\r'."""
    starts = []
    running_index = 0
    for text_line in lines_for_write:
        starts.append(running_index)
        # each paragraph ends with '\r'
        running_index += len(text_line) + 1
    return starts


def apply_formatting_to_textnote(textnote, lines_for_write, meta_per_line):
    """Write content and apply list formatting per paragraph."""
    text_with_lf = u"\n".join(lines_for_write)
    text_with_cr = denormalize_to_revit(text_with_lf)
    textnote.Text = text_with_cr

    starts = map_line_starts(lines_for_write)
    fmt = textnote.GetFormattedText()

    count = len(lines_for_write)
    for index in range(count):
        level = meta_per_line[index]["level"]
        if level is None:
            continue

        start_char = starts[index]
        length = len(lines_for_write[index]) + 1  # include trailing '\r'
        text_range = TextRange(start_char, length)

        try:
            fmt.SetListType(text_range, level.list_type)
            fmt.SetIndentLevel(text_range, level.indent)
        except Exception as ex:
            LOGGER.warning("List format failed at paragraph {}: {}".format(index, ex))

    textnote.SetFormattedText(fmt)


def main():
    textnote = pick_single_textnote()
    if not textnote:
        forms.alert("Please select a single Text Note.", exitscript=True)
        return

    source_text = normalize_newlines(textnote.Text or u"")
    source_lines = source_text.split(u"\n")

    # 1) Split by numeric section headers (each section resets lists)
    sections = split_into_sections(source_lines)

    # 2) Build paragraphs and metadata (indents and list levels); no '\v' yet
    final_lines, meta_per_line = rebuild_text_and_meta(sections)

    # 3) Add '\v' between list items only if the level's blank_between is True
    augmented_lines = add_vertical_tabs_at_list_runs(final_lines, meta_per_line)

    # 4) Write and format against the augmented lines
    with Transaction(revit.doc, "Format TextNote Lists (configurable)") as tx:
        tx.Start()
        apply_formatting_to_textnote(textnote, augmented_lines, meta_per_line)
        tx.Commit()

    forms.alert("Done. Lists formatted; indicators removed; list tabs cleaned; section breaks spaced.", ok=True)


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        LOGGER.exception("Formatting failed: {}".format(ex))
        forms.alert("Formatting failed:\n{}".format(ex), ok=True)
