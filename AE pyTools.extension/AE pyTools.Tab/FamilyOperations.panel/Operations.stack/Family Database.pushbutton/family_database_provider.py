# -*- coding: utf-8 -*-
"""Backend providers for Family Database UI.

To migrate from YAML to SQL later, keep the UI unchanged and implement the SQL
query logic in ``SqlFamilyDatabaseProvider`` using the same method contract:
    - list_entries()
    - get_type_rows(entry)
"""

import os

from pyrevit.coreutils import yaml


DEFAULT_CATEGORY = 'Uncategorized'


try:
    basestring
except NameError:
    basestring = str


def _to_text(value):
    if value is None:
        return ''
    try:
        return unicode(value)
    except Exception:
        return str(value)


def _clean_text(value):
    return _to_text(value).strip()


def _normalize_key(value):
    key = _clean_text(value).lower().replace('_', ' ').replace('-', ' ')
    return ' '.join([part for part in key.split() if part])


class FamilyDatabaseEntry(object):
    def __init__(
        self,
        key,
        family_name='',
        client_name='',
        project_number='',
        category=DEFAULT_CATEGORY,
        source_label='',
        file_name='',
    ):
        self.key = key
        self.family_name = _clean_text(family_name)
        self.client_name = _clean_text(client_name)
        self.project_number = _clean_text(project_number)
        self.category = _clean_text(category) or DEFAULT_CATEGORY
        self.source_label = _clean_text(source_label)
        self.file_name = _clean_text(file_name)

        suffix_tokens = []
        if self.client_name:
            suffix_tokens.append(self.client_name)
        if self.project_number:
            suffix_tokens.append(self.project_number)
        suffix = ' | '.join(suffix_tokens)

        if suffix:
            self.label = '{} ({})'.format(self.family_name, suffix)
        else:
            self.label = self.family_name

    def matches(self, family_filter, client_filter, project_filter, selected_category):
        if selected_category and selected_category != 'All Categories':
            if self.category != selected_category:
                return False

        family_filter = _clean_text(family_filter).lower()
        client_filter = _clean_text(client_filter).lower()
        project_filter = _clean_text(project_filter).lower()

        if family_filter:
            family_haystack = '{} {}'.format(self.family_name, self.file_name).lower()
            if family_filter not in family_haystack:
                return False

        if client_filter and client_filter not in self.client_name.lower():
            return False

        if project_filter and project_filter not in self.project_number.lower():
            return False

        return True

    def __str__(self):
        return self.label


class FamilyDatabaseProviderBase(object):
    backend_name = 'base'

    def list_entries(self):
        raise NotImplementedError

    def get_type_rows(self, entry):
        raise NotImplementedError

    def backend_summary(self):
        return self.backend_name.upper()


class YamlFamilyDatabaseProvider(FamilyDatabaseProviderBase):
    backend_name = 'yaml'

    def __init__(self, source_dir, map_file_name='map.yaml'):
        self._source_dir = source_dir
        self._map_file_name = (map_file_name or 'map.yaml').lower()

    def _parse_filename_tokens(self, yaml_path):
        stem = os.path.splitext(os.path.basename(yaml_path))[0]
        parts = stem.rsplit('_', 2)
        if len(parts) == 3:
            return parts[0], parts[1], parts[2]
        return stem, '', ''

    def _category_for_path(self, yaml_path):
        try:
            rel_parent = os.path.relpath(os.path.dirname(yaml_path), self._source_dir)
        except Exception:
            rel_parent = ''

        rel_parent = _clean_text(rel_parent)
        if not rel_parent or rel_parent == '.' or rel_parent.startswith('..'):
            return DEFAULT_CATEGORY

        return rel_parent.split(os.sep)[0]

    def _extract_product_type(self, type_payload):
        if not isinstance(type_payload, dict):
            return ''
        for param_name, param_value in type_payload.items():
            normalized = _normalize_key(param_name)
            if normalized in ('product type', 'producttype'):
                return _clean_text(param_value)
        return ''

    def backend_summary(self):
        return 'YAML: {}'.format(self._source_dir)

    def list_entries(self):
        if not os.path.isdir(self._source_dir):
            return []

        entries = []
        for walk_root, _, files in os.walk(self._source_dir):
            for file_name in files:
                file_lower = file_name.lower()
                if not file_lower.endswith('.yaml'):
                    continue
                if file_lower == self._map_file_name:
                    continue

                yaml_path = os.path.join(walk_root, file_name)
                family_token, client_name, project_number = self._parse_filename_tokens(yaml_path)
                family_name = family_token.replace('_', ' ').strip() or family_token

                entries.append(
                    FamilyDatabaseEntry(
                        key=yaml_path,
                        family_name=family_name,
                        client_name=client_name,
                        project_number=project_number,
                        category=self._category_for_path(yaml_path),
                        source_label=yaml_path,
                        file_name=file_name,
                    )
                )

        entries = sorted(
            entries,
            key=lambda e: (
                e.category.lower(),
                e.family_name.lower(),
                e.client_name.lower(),
                e.project_number.lower(),
                e.file_name.lower(),
            )
        )
        return entries

    def get_type_rows(self, entry):
        loaded = yaml.load_as_dict(entry.key)
        if not isinstance(loaded, dict):
            return []

        type_rows = []
        type_data = loaded.get('types', {}) or {}
        if isinstance(type_data, dict):
            for type_name, type_payload in type_data.items():
                type_rows.append({
                    'type_name': _clean_text(type_name),
                    'product_type': self._extract_product_type(type_payload),
                    'client_name': entry.client_name,
                    'project_number': entry.project_number,
                })

        return sorted(type_rows, key=lambda row: row['type_name'].lower())


class SqlFamilyDatabaseProvider(FamilyDatabaseProviderBase):
    backend_name = 'sql'

    def __init__(self, connection_string):
        self._connection_string = _clean_text(connection_string)

    def backend_summary(self):
        if self._connection_string:
            return 'SQL: configured connection'
        return 'SQL: no connection configured'

    def list_entries(self):
        # Intentionally returns no rows until SQL schema/query contract is set.
        # Keep the method signature stable so UI can switch backends without edits.
        return []

    def get_type_rows(self, entry):
        return []


def create_family_database_provider(
    backend_name,
    source_dir=None,
    connection_string=None,
    map_file_name='map.yaml'
):
    backend = _clean_text(backend_name).lower() or 'yaml'

    if backend == 'yaml':
        return YamlFamilyDatabaseProvider(
            source_dir=source_dir,
            map_file_name=map_file_name
        )

    if backend == 'sql':
        return SqlFamilyDatabaseProvider(connection_string=connection_string)

    raise ValueError('Unsupported family database backend: {}'.format(backend_name))
