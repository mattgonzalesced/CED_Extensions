# -*- coding: utf-8 -*-
"""Family Database UI for browsing and running family import/export workflows."""

import os

from pyrevit import forms
from pyrevit import script

from family_database_provider import create_family_database_provider


logger = script.get_logger()

WINDOW_XAML = os.path.join(os.path.dirname(__file__), 'window.xaml')

FAMILY_CONFIGS_ROOT = r'C:\ACC\ACCDocs\CoolSys\CED Content Collection\Project Files\03 Automations\Family Configs'
FAMILY_CONFIGS_SOURCE = os.path.join(FAMILY_CONFIGS_ROOT, 'Source')
MAP_FILE_NAME = 'map.yaml'

# Switch backend by setting CED_FAMILY_DB_BACKEND to "yaml" or "sql".
DATA_BACKEND = os.environ.get('CED_FAMILY_DB_BACKEND', 'yaml')
SQL_CONNECTION_STRING = os.environ.get('CED_FAMILY_DB_CONNECTION', '')

OPERATIONS_STACK_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
IMPORT_COMMAND_SCRIPT = os.path.join(
    OPERATIONS_STACK_DIR,
    'Import-Export.pulldown',
    'ImportFamilyTypes.pushbutton',
    'script.py'
)
EXPORT_COMMAND_SCRIPT = os.path.join(
    OPERATIONS_STACK_DIR,
    'Import-Export.pulldown',
    'ExportFamilyTypes.pushbutton',
    'script.py'
)


def _to_text(value):
    if value is None:
        return ''
    try:
        return unicode(value)
    except Exception:
        return str(value)


def _clean_text(value):
    return _to_text(value).strip()


def _execute_script_file(script_path):
    try:
        import runpy
    except Exception:
        runpy = None

    if runpy:
        runpy.run_path(script_path, run_name='__main__')
        return

    namespace = {'__name__': '__main__', '__file__': script_path}
    try:
        _execfile = execfile
    except NameError:
        _execfile = None

    if _execfile:
        _execfile(script_path, namespace, namespace)
        return

    raise Exception('Could not execute script using runpy or execfile: {}'.format(script_path))


class FamilyDatabaseWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, WINDOW_XAML)

        self._provider = create_family_database_provider(
            backend_name=DATA_BACKEND,
            source_dir=FAMILY_CONFIGS_SOURCE,
            connection_string=SQL_CONNECTION_STRING,
            map_file_name=MAP_FILE_NAME
        )
        self._all_entries = []
        self._filtered_entries = []

        self._category_combo = self.FindName('CategoryCombo')
        self._family_filter_box = self.FindName('FamilyFilterBox')
        self._client_filter_box = self.FindName('ClientFilterBox')
        self._project_filter_box = self.FindName('ProjectFilterBox')
        self._family_files_list = self.FindName('FamilyFilesList')
        self._type_rows_list = self.FindName('TypeRowsList')
        self._selected_file_text = self.FindName('SelectedFileText')
        self._status_text = self.FindName('StatusText')
        self._import_button = self.FindName('ImportButton')
        self._export_button = self.FindName('ExportButton')
        self._refresh_button = self.FindName('RefreshButton')
        self._close_button = self.FindName('CloseButton')

        self._wire_events()
        self._load_index()

    def _wire_events(self):
        if self._category_combo:
            self._category_combo.SelectionChanged += self._on_filter_changed
        if self._family_filter_box:
            self._family_filter_box.TextChanged += self._on_filter_changed
        if self._client_filter_box:
            self._client_filter_box.TextChanged += self._on_filter_changed
        if self._project_filter_box:
            self._project_filter_box.TextChanged += self._on_filter_changed
        if self._family_files_list:
            self._family_files_list.SelectionChanged += self._on_family_selected
        if self._import_button:
            self._import_button.Click += self._on_import_clicked
        if self._export_button:
            self._export_button.Click += self._on_export_clicked
        if self._refresh_button:
            self._refresh_button.Click += self._on_refresh_clicked
        if self._close_button:
            self._close_button.Click += self._on_close_clicked

    def _current_category(self):
        if not self._category_combo:
            return 'All Categories'
        selection = self._category_combo.SelectedItem
        return _clean_text(selection) or 'All Categories'

    def _textbox_value(self, textbox):
        if textbox is None:
            return ''
        return _clean_text(getattr(textbox, 'Text', ''))

    def _set_status(self, text):
        if self._status_text is not None:
            self._status_text.Text = _clean_text(text)

    def _set_selected_file_text(self, text):
        if self._selected_file_text is not None:
            self._selected_file_text.Text = _clean_text(text)

    def _find_selected_entry(self):
        if not self._family_files_list:
            return None
        selected_item = self._family_files_list.SelectedItem
        if selected_item is None:
            return None
        return selected_item

    def _load_index(self):
        prior_selected = self._find_selected_entry()
        prior_selected_key = getattr(prior_selected, 'key', None)
        prior_category = self._current_category()

        try:
            self._all_entries = self._provider.list_entries()
        except Exception as ex:
            logger.exception('Failed to load family database entries.')
            forms.alert(
                'Failed to load family database entries:\n{}'.format(ex),
                title='Family Database',
                warn_icon=True
            )
            self._all_entries = []

        categories = sorted(
            set([entry.category for entry in self._all_entries]),
            key=lambda c: c.lower()
        )

        category_items = ['All Categories'] + categories
        if self._category_combo is not None:
            self._category_combo.ItemsSource = category_items
            if prior_category in category_items:
                self._category_combo.SelectedItem = prior_category
            else:
                self._category_combo.SelectedIndex = 0

        self._apply_filters(selected_key_hint=prior_selected_key)

    def _apply_filters(self, selected_key_hint=None):
        selected_category = self._current_category()
        family_filter = self._textbox_value(self._family_filter_box)
        client_filter = self._textbox_value(self._client_filter_box)
        project_filter = self._textbox_value(self._project_filter_box)

        filtered = []
        for entry in self._all_entries:
            if entry.matches(
                family_filter=family_filter,
                client_filter=client_filter,
                project_filter=project_filter,
                selected_category=selected_category
            ):
                filtered.append(entry)
        self._filtered_entries = filtered

        if self._family_files_list is not None:
            self._family_files_list.ItemsSource = filtered

        if not filtered:
            if self._family_files_list is not None:
                self._family_files_list.SelectedItem = None
            self._type_rows_list.ItemsSource = []
            self._set_selected_file_text('No family file selected.')
            self._set_status(
                '{} | 0 files found'.format(self._provider.backend_summary())
            )
            return

        selected_entry = None
        if selected_key_hint:
            for entry in filtered:
                if _clean_text(entry.key).lower() == _clean_text(selected_key_hint).lower():
                    selected_entry = entry
                    break

        if selected_entry is None and self._family_files_list is not None:
            current = self._family_files_list.SelectedItem
            if current in filtered:
                selected_entry = current

        if selected_entry is None:
            selected_entry = filtered[0]

        if self._family_files_list is not None:
            self._family_files_list.SelectedItem = selected_entry

        self._load_type_rows(selected_entry)

    def _load_type_rows(self, entry):
        if entry is None:
            self._type_rows_list.ItemsSource = []
            self._set_selected_file_text('No family file selected.')
            return

        self._set_selected_file_text(entry.source_label or entry.key)
        try:
            type_rows = self._provider.get_type_rows(entry)
        except Exception as ex:
            logger.exception('Failed to load type rows for entry: %s', entry.key)
            forms.alert(
                'Failed to load family types:\n{}\n\n{}'.format(entry.key, ex),
                title='Family Database',
                warn_icon=True
            )
            type_rows = []

        self._type_rows_list.ItemsSource = type_rows
        self._set_status(
            '{} | {} files shown | {} types in selected file'.format(
                self._provider.backend_summary(),
                len(self._filtered_entries),
                len(type_rows)
            )
        )

    def _run_existing_command(self, command_script_path, friendly_name):
        if not os.path.isfile(command_script_path):
            forms.alert(
                '{} script not found:\n{}'.format(friendly_name, command_script_path),
                title='Family Database',
                warn_icon=True
            )
            return

        try:
            self.Hide()
        except Exception:
            pass

        try:
            _execute_script_file(command_script_path)
        except SystemExit:
            pass
        except Exception as ex:
            logger.exception('Failed running %s command script.', friendly_name)
            forms.alert(
                '{} failed:\n{}'.format(friendly_name, ex),
                title='Family Database',
                warn_icon=True
            )
        finally:
            try:
                self.Show()
                self.Activate()
            except Exception:
                pass
            self._load_index()

    def _on_filter_changed(self, sender, args):
        selected = self._find_selected_entry()
        selected_key = getattr(selected, 'key', None)
        self._apply_filters(selected_key_hint=selected_key)

    def _on_family_selected(self, sender, args):
        selected = self._find_selected_entry()
        self._load_type_rows(selected)

    def _on_import_clicked(self, sender, args):
        self._run_existing_command(IMPORT_COMMAND_SCRIPT, 'Import Family')

    def _on_export_clicked(self, sender, args):
        self._run_existing_command(EXPORT_COMMAND_SCRIPT, 'Export Family')

    def _on_refresh_clicked(self, sender, args):
        self._load_index()

    def _on_close_clicked(self, sender, args):
        self.Close()


if __name__ == '__main__':
    window = FamilyDatabaseWindow()
    window.ShowDialog()
