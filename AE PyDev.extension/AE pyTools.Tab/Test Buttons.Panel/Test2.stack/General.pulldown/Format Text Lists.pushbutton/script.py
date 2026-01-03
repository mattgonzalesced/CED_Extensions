# -*- coding: utf-8 -*-
# IronPython 2.7 / pyRevit / Revit API
#
# UI-driven TextNote list formatter (5 rows configurable)
# - PatternType detection (source prefixes)
# - Replace with Revit list type + indent
# - Optional section header split/reset (spec mode)
# - Removes true blank lines
# - Merges indented hard-wrap continuation lines into previous list item using '\v'
# - Optional '\v' between list items (blank_between flag)

import re

from Autodesk.Revit.DB import Transaction, TextRange, ListType
from pyrevit import revit, DB, forms, script
from pyrevit.forms import Reactive, reactive

LOGGER = script.get_logger()

# e.g. "15100.02 INTENT" or "3.01  INSTALLATION" or "1.07 SUBMITTALS"
RE_SECTION_HEADER = re.compile(r'^\s*\d{1,5}\.\d+(\s|$)')


class PatternType(object):
    UPPER = "UPPER"              # A. B. C. AA. AB.
    LOWER = "LOWER"              # a. b. c.
    NUMERIC = "NUMERIC"          # 1. 2. 3.
    DECIMAL = "DECIMAL"          # 1.1 1.2 2.1
    TRIDECIMAL = "TRIDECIMAL"    # 1.1.1 1.1.2 2.3.4
    NUMDOT = "NUMDOT"            # 1.A. 1.a. 2.AA.


PATTERN_EXAMPLES = {
    PatternType.NUMERIC: "1. 2. 3.",
    PatternType.DECIMAL: "1.1 1.2 2.1",
    PatternType.TRIDECIMAL: "1.1.1 1.1.2 2.3.4",
    PatternType.UPPER: "A. B. C. AA.",
    PatternType.LOWER: "a. b. c.",
    PatternType.NUMDOT: "1.A. 2.B. 10.AA.",
}


REVIT_LIST_KEYS = [
    "Bullet",
    "ArabicNumbers",
    "LowerCaseLetters",
    "UpperCaseLetters",
]

REVIT_LIST_EXAMPLES = {
    "Bullet": u"• • •",
    "ArabicNumbers": "1. 2. 3.",
    "LowerCaseLetters": "a. b. c.",
    "UpperCaseLetters": "A. B. C.",
}

REVIT_LIST_KEY_TO_ENUM = {
    "Bullet": ListType.Bullet,
    "ArabicNumbers": ListType.ArabicNumbers,
    "LowerCaseLetters": ListType.LowerCaseLetters,
    "UpperCaseLetters": ListType.UpperCaseLetters,
}



class OptionItem(object):
    def __init__(self, key, display):
        self.Key = key
        self.Display = display

class TokenType(object):
    DOT = "DOT"        # 1.
    RPAREN = "RPAREN"  # 1)
    DASH = "DASH"      # 1 -

class LevelRow(Reactive):
    def __init__(self, level_index):
        Reactive.__init__(self)
        self._level_index = level_index

        self._enabled = False

        self._pattern_key = ""
        self._token_key = ""
        self._revit_list_key = ""
        self._indent = ""  # keep blank until enabled

        self._blank_before = False
        self._blank_between = False
        self._blank_after = False

    @reactive
    def LevelIndex(self):
        return self._level_index

    @reactive
    def Enabled(self):
        return self._enabled

    @Enabled.setter
    def Enabled(self, value):
        self._enabled = bool(value)

    @reactive
    def PatternKey(self):
        return self._pattern_key

    @PatternKey.setter
    def PatternKey(self, value):
        self._pattern_key = value or ""

    @reactive
    def TokenKey(self):
        return self._token_key

    @TokenKey.setter
    def TokenKey(self, value):
        self._token_key = value or ""

    @reactive
    def RevitListKey(self):
        return self._revit_list_key

    @RevitListKey.setter
    def RevitListKey(self, value):
        self._revit_list_key = value or ""

    @reactive
    def Indent(self):
        return self._indent

    @Indent.setter
    def Indent(self, value):
        if value is None:
            self._indent = ""
        else:
            self._indent = str(value)

    @reactive
    def BlankBefore(self):
        return self._blank_before

    @BlankBefore.setter
    def BlankBefore(self, value):
        self._blank_before = bool(value)

    @reactive
    def BlankBetween(self):
        return self._blank_between

    @BlankBetween.setter
    def BlankBetween(self, value):
        self._blank_between = bool(value)

    @reactive
    def BlankAfter(self):
        return self._blank_after

    @BlankAfter.setter
    def BlankAfter(self, value):
        self._blank_after = bool(value)

    def clear_row(self):
        self.PatternKey = ""
        self.TokenKey = ""
        self.RevitListKey = ""
        self.Indent = ""
        self.BlankBefore = False
        self.BlankBetween = False
        self.BlankAfter = False

    def apply_defaults_for_level(self):
        self.Indent = str(self.LevelIndex - 1)  # 0..4

        if self.LevelIndex == 1:
            self.PatternKey = PatternType.NUMERIC
            self.TokenKey = TokenType.DOT
            self.RevitListKey = "ArabicNumbers"
        elif self.LevelIndex == 2:
            self.PatternKey = PatternType.LOWER
            self.TokenKey = TokenType.DOT
            self.RevitListKey = "LowerCaseLetters"
        elif self.LevelIndex == 3:
            self.PatternKey = PatternType.UPPER
            self.TokenKey = TokenType.DOT
            self.RevitListKey = "UpperCaseLetters"
        else:
            self.PatternKey = PatternType.NUMERIC
            self.TokenKey = TokenType.DOT
            self.RevitListKey = "ArabicNumbers"







class ListLevel(object):
    def __init__(self, pattern_type, list_type, indent,
                 blank_before=False, blank_between=False, blank_after=False,
                 token_type=None):
        self.pattern_type = pattern_type
        self.list_type = list_type
        self.indent = indent
        self.blank_before = blank_before
        self.blank_between = blank_between
        self.blank_after = blank_after
        self.token_type = token_type or TokenType.DOT

        if self.token_type == TokenType.RPAREN:
            token_re = r'\)'
            tail_re = r'(?:\s*|$)'
        elif self.token_type == TokenType.DASH:
            token_re = r'\s*-\s*'
            tail_re = r'(?:\s*|$)'
        else:
            token_re = r'\.'
            tail_re = r'(?:\s+|$)'   # IMPORTANT: keep strict for DOT


        if pattern_type == PatternType.UPPER:
            self.regex = re.compile(r'^\s*([A-Z]{1,3})' + token_re + tail_re)
        elif pattern_type == PatternType.LOWER:
            self.regex = re.compile(r'^\s*([a-z]{1,3})' + token_re + tail_re)
        elif pattern_type == PatternType.NUMERIC:
            self.regex = re.compile(r'^\s*(\d+)\.(?!\d)(?:\s+|$)')
        elif pattern_type == PatternType.DECIMAL:
            # 17.1. or 17.1  (but NOT 17.1.2)
            self.regex = re.compile(r'^\s*(\d+)\.(\d+)(?:\.(?!\d)|\s+|$)')
        elif pattern_type == PatternType.TRIDECIMAL:
            # 17.1.2. or 17.1.2 (but NOT 17.1.2.3)
            self.regex = re.compile(r'^\s*(\d+)\.(\d+)\.(\d+)(?:\.(?!\d)|\s+|$)')

        elif pattern_type == PatternType.NUMDOT:
            self.regex = re.compile(r'^\s*\d+\.\s*([A-Za-z]{1,3})\.(?:\s+|$)')
        else:
            self.regex = re.compile(r'^(?!)')


    def match(self, text):
        return self.regex.match(text)


class ListConfigWindow(forms.WPFWindow):
    def __init__(self, xaml_path):
        # Gate all UI events until the window is loaded
        self._ui_ready = False

        forms.WPFWindow.__init__(self, xaml_path)

        # dropdown options
        self.PatternOptions = []
        for key in [PatternType.NUMERIC, PatternType.DECIMAL, PatternType.TRIDECIMAL,
                    PatternType.NUMDOT, PatternType.UPPER, PatternType.LOWER]:
            self.PatternOptions.append(OptionItem(key, key))

        self.RevitListOptions = []
        for key in REVIT_LIST_KEYS:
            self.RevitListOptions.append(OptionItem(key, key))

        self.TokenOptions = []
        for key in [TokenType.DOT, TokenType.RPAREN, TokenType.DASH]:
            self.TokenOptions.append(OptionItem(key, key))

        # 5 rows
        self.Levels = []
        for i in range(1, 6):
            self.Levels.append(LevelRow(i))

        # enable first two by default
        self.Levels[0].Enabled = True
        self.Levels[0].apply_defaults_for_level()

        self.Levels[1].Enabled = True
        self.Levels[1].apply_defaults_for_level()

        # remaining disabled + cleared
        for idx in [2, 3, 4]:
            self.Levels[idx].Enabled = False
            self.Levels[idx].clear_row()

        self.ChkOperateExisting.IsChecked = False

        # bind window context
        self.DataContext = self

        # spec mode default off
        self.ChkSectionHeaders.IsChecked = False

        # DO NOT autofix here (causes event storm during grid load)

        # wiring
        self.BtnCancel.Click += self._on_cancel
        self.BtnRun.Click += self._on_run

        # Mark UI ready only after window is loaded and grid is created
        try:
            self.Loaded += self._on_loaded
        except Exception:
            # Some hosts don’t expose Loaded; if so, we’ll just enable immediately
            self._ui_ready = True

        self._ok = False

    def _on_loaded(self, sender, args):
        # Now it’s safe to normalize once
        try:
            self._autofix_enabled_rows()
        except Exception:
            pass

        self._ui_ready = True


    def _on_cancel(self, sender, args):
        self._ok = False
        self.Close()

    def _on_run(self, sender, args):
        try:
            # Commit any pending edits in the grid (combobox/checkbox/textbox)
            self.GridLevels.CommitEdit()
            self.GridLevels.CommitEdit()
        except Exception:
            pass

        try:
            self._autofix_enabled_rows()
        except Exception:
            pass

        self._ok = True
        self.Close()

    @property
    def ok(self):
        return self._ok

    @property
    def enable_section_headers(self):
        try:
            return bool(self.ChkSectionHeaders.IsChecked)
        except Exception:
            return False

    @property
    def operate_existing(self):
        try:
            return bool(self.ChkOperateExisting.IsChecked)
        except Exception:
            return False

    def _autofix_enabled_rows(self):
        """
        - Disabled rows are cleared (blank display)
        - Enabled rows get defaults if empty
        - Enabled indent values are enforced to be strictly increasing (>= prev + 1)
        - If you bump level 1 to 1, level 2+ auto becomes 2,3,4...
        """
        # 1) Clear disabled / ensure defaults for enabled
        for row in self.Levels:
            if not row.Enabled:
                row.clear_row()
                continue

            # if enabled but values missing, restore defaults
            if not row.PatternKey or not row.RevitListKey:
                row.apply_defaults_for_level()

            if not row.TokenKey:
                row.TokenKey = TokenType.DOT

            if row.Indent in (None, ""):
                row.Indent = str(row.LevelIndex - 1)

        # 2) Enforce indents increasing across enabled rows
        enabled = [r for r in self.Levels if r.Enabled]
        enabled.sort(key=lambda x: x.LevelIndex)

        prev = None
        for r in enabled:
            try:
                cur = int(r.Indent)
            except Exception:
                cur = r.LevelIndex - 1

            if cur < 0:
                cur = 0

            if prev is not None and cur <= prev:
                cur = prev + 1

            r.Indent = str(cur)
            prev = cur

    # These two are called by XAML event hooks (I’m giving you the handlers now)
    def EnabledChanged(self, sender, args):
        if not self._ui_ready:
            return

        row = getattr(sender, "DataContext", None)
        if not row:
            return

        if row.Enabled:
            row.apply_defaults_for_level()
        else:
            row.clear_row()

        self._autofix_enabled_rows()
        # NO GridLevels.Items.Refresh() — do not force rebuilds

    def IndentLostFocus(self, sender, args):
        if not self._ui_ready:
            return

        self._autofix_enabled_rows()
        # NO GridLevels.Items.Refresh()

def add_spacing_soft_returns_only(textnote):
    """
    Adds a soft return (\v) at the end of each non-empty paragraph.
    Does NOT attempt to detect list vs non-list. Fast + reliable.
    """
    src = normalize_newlines(textnote.Text or u"")
    lines = src.split(u"\n")

    out_lines = []
    for i in range(len(lines)):
        ln = lines[i]
        if ln.strip() == u"":
            out_lines.append(ln)
            continue

        # already has trailing \v? leave it
        if ln.endswith(u"\v"):
            out_lines.append(ln)
        else:
            out_lines.append(ln + u"\v")

    # write back text (this may reset formatting in some cases)
    # so preserve existing formatted text object and re-apply it after setting text
    fmt = textnote.GetFormattedText()
    textnote.Text = denormalize_to_revit(u"\n".join(out_lines))
    textnote.SetFormattedText(fmt)


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
    if not line_text:
        return line_text

    has_trailing_vt = line_text.endswith(u"\v")
    if has_trailing_vt:
        core = line_text[:-1]
    else:
        core = line_text

    core = re.sub(u'[ \t]+', u' ', core)
    core = core.replace(u"\v", u"")

    if has_trailing_vt:
        return core + u"\v"
    return core


def is_indented_continuation(raw_line):
    if not raw_line:
        return False
    return raw_line.startswith(u"\t") or raw_line.startswith(u" ") or raw_line.startswith(u"\u00A0")


def strip_leading_whitespace(line_text):
    return re.sub(u'^[ \t\u00A0]+', u'', line_text)


def build_list_levels_from_rows(rows):
    # Keep only enabled rows, in level order (1..5)
    enabled_rows = [r for r in rows if r.Enabled]
    enabled_rows.sort(key=lambda x: x.LevelIndex)

    if not enabled_rows:
        return []

    # Parse user indents and enforce: indent[i] >= indent[i-1] + 1
    parsed = []
    for r in enabled_rows:
        try:
            val = int(r.Indent)
        except Exception:
            val = r.LevelIndex - 1

        if val < 0:
            val = 0

        parsed.append(val)

    enforced = []
    for i in range(len(parsed)):
        if i == 0:
            enforced.append(parsed[i])
        else:
            prev = enforced[i - 1]
            cur = parsed[i]

            # Preserve gaps if user chose a larger indent,
            # but enforce at least +1 deeper than previous.
            if cur <= prev:
                cur = prev + 1

            enforced.append(cur)

    # Build ListLevel objects using enforced indents
    levels = []
    for i in range(len(enabled_rows)):
        row = enabled_rows[i]

        list_enum = REVIT_LIST_KEY_TO_ENUM.get(row.RevitListKey, None)
        if list_enum is None:
            list_enum = ListType.ArabicNumbers

        # Optional: log if we changed the indent
        if enforced[i] != parsed[i]:
            LOGGER.debug("Indent auto-fix: Level {} from {} to {}".format(
                row.LevelIndex, parsed[i], enforced[i]
            ))

        levels.append(
            ListLevel(
                row.PatternKey,
                list_type=list_enum,
                indent=enforced[i],
                blank_before=bool(row.BlankBefore),
                blank_between=bool(row.BlankBetween),
                blank_after=bool(row.BlankAfter),
                token_type=row.TokenKey
            )
        )


    return levels



def split_into_sections(lines, enable_section_headers):
    if not enable_section_headers:
        return [{"header": None, "body": list(lines)}]

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


def classify_and_strip(line_text, list_levels):
    for level in list_levels:
        match_obj = level.match(line_text)
        if match_obj:
            stripped = line_text[match_obj.end():]
            stripped = clean_internal_tabs(stripped)
            return level, stripped
    return None, line_text


def rebuild_text_and_meta(sections, list_levels, enable_section_headers):
    final_lines = []
    meta_per_line = []
    last_level = None

    for section in sections:
        header = section["header"]
        body_lines = section["body"]

        # header paragraph resets list run
        if header is not None:
            header_text = header.strip()
            if header_text != u"":
                if last_level is not None and len(final_lines) > 0:
                    final_lines.append(u"")
                    meta_per_line.append({"level": None})
                    last_level = None

                final_lines.append(header_text)
                meta_per_line.append({"level": None})
                last_level = None

        for raw_line in body_lines:
            if raw_line.strip() == u"":
                continue

            # merge hard-wrapped continuations into previous list item
            if last_level is not None and is_indented_continuation(raw_line):
                peek_level, _ = classify_and_strip(raw_line, list_levels)
                if peek_level is None and (not enable_section_headers or not RE_SECTION_HEADER.match(raw_line)):
                    cont = strip_leading_whitespace(raw_line)
                    cont = clean_internal_tabs(cont)
                    if cont.strip() != u"":
                        final_lines[-1] = final_lines[-1] + u"\v" + cont
                    continue

            level, content = classify_and_strip(raw_line, list_levels)

            if level is not None and content.strip() == u"":
                continue

            final_lines.append(content)
            meta_per_line.append({"level": level})
            last_level = level

    return final_lines, meta_per_line


def add_vertical_tabs_at_list_runs(final_lines, meta_per_line):
    augmented = []
    total = len(final_lines)

    for i in range(total):
        line = final_lines[i]
        level = meta_per_line[i]["level"]

        if level is not None:
            if i < total - 1:
                next_level = meta_per_line[i + 1]["level"]
            else:
                next_level = None

            line = clean_internal_tabs(line)

            if level.blank_between and next_level is not None:
                if not line.endswith(u"\v"):
                    line = line + u"\v"

        augmented.append(line)

    return augmented


def map_line_starts(lines_for_write):
    starts = []
    running = 0
    for ln in lines_for_write:
        starts.append(running)
        running += len(ln) + 1  # +1 for '\r'
    return starts


def apply_formatting_to_textnote(textnote, lines_for_write, meta_per_line):
    # Write full text
    text_with_lf = u"\n".join(lines_for_write)
    textnote.Text = denormalize_to_revit(text_with_lf)

    starts = map_line_starts(lines_for_write)
    fmt = textnote.GetFormattedText()

    # Collect list paragraphs
    items = []
    for index in range(len(lines_for_write)):
        level = meta_per_line[index]["level"]
        if level is None:
            continue

        start_char = starts[index]
        length = len(lines_for_write[index])
        if length <= 0:
            continue

        items.append({
            "index": index,
            "level": level,
            "start": start_char,
            "length": length,
            "indent": level.indent,
        })

    if not items:
        textnote.SetFormattedText(fmt)
        return

    # Determine top-level indent (minimum indent among list items)
    min_indent = None
    for it in items:
        if min_indent is None or it["indent"] < min_indent:
            min_indent = it["indent"]

    # Pass 1: apply indents (order doesn't matter much)
    for it in items:
        tr = TextRange(it["start"], it["length"])
        try:
            fmt.SetIndentLevel(tr, it["indent"])
        except Exception as ex:
            LOGGER.warning("SetIndentLevel failed at paragraph {}: {}".format(it["index"], ex))

    # Pass 2: apply list types deepest -> shallowest
    # (This avoids deeper levels overriding earlier top-level list templates)
    items_by_indent = sorted(items, key=lambda x: x["indent"], reverse=True)
    for it in items_by_indent:
        tr = TextRange(it["start"], it["length"])
        try:
            fmt.SetListType(tr, it["level"].list_type)
        except Exception as ex:
            LOGGER.warning("SetListType failed at paragraph {}: {}".format(it["index"], ex))

    # Pass 3: reapply top-level list type (min indent) to ensure it sticks
    for it in items:
        if it["indent"] != min_indent:
            continue
        tr = TextRange(it["start"], it["length"])
        try:
            fmt.SetListType(tr, it["level"].list_type)
        except Exception as ex:
            LOGGER.warning("Top-level SetListType reapply failed at paragraph {}: {}".format(it["index"], ex))

    textnote.SetFormattedText(fmt)



def main():
    textnote = pick_single_textnote()
    if not textnote:
        forms.alert("Please select a single Text Note.", exitscript=True)
        return

    xaml_path = script.get_bundle_file("ListConfigWindow.xaml")
    dlg = ListConfigWindow(xaml_path)
    dlg.ShowDialog()

    if not dlg.ok:
        return

    enable_section_headers = dlg.enable_section_headers
    list_levels = build_list_levels_from_rows(dlg.Levels)

    if not list_levels:
        forms.alert("No enabled rows. Nothing to do.", ok=True)
        return

    source_text = normalize_newlines(textnote.Text or u"")
    source_lines = source_text.split(u"\n")

    sections = split_into_sections(source_lines, enable_section_headers)
    final_lines, meta_per_line = rebuild_text_and_meta(sections, list_levels, enable_section_headers)
    augmented_lines = add_vertical_tabs_at_list_runs(final_lines, meta_per_line)

    operate_existing = dlg.operate_existing

    with Transaction(revit.doc, "Format TextNote Lists (UI)") as tx:
        tx.Start()

        if operate_existing:
            # spacing only, no rebuild
            add_spacing_soft_returns_only(textnote)
        else:
            # your existing rebuild pipeline
            sections = split_into_sections(source_lines, enable_section_headers)
            final_lines, meta_per_line = rebuild_text_and_meta(sections, list_levels, enable_section_headers)
            augmented_lines = add_vertical_tabs_at_list_runs(final_lines, meta_per_line)
            apply_formatting_to_textnote(textnote, augmented_lines, meta_per_line)

        tx.Commit()

    forms.alert("Done.", ok=True)


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        LOGGER.exception("Formatting failed: {}".format(ex))
        forms.alert("Formatting failed:\n{}".format(ex), ok=True)
