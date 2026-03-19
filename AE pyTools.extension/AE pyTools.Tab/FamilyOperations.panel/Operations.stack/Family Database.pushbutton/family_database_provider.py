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
        self.display_name = self.label

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

    def __repr__(self):
        return self.label

    def ToString(self):
        # WPF ListBox falls back to .NET ToString() for plain item display.
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
        self._map_categories_by_file = None

    def _parse_filename_tokens(self, yaml_path):
        stem = os.path.splitext(os.path.basename(yaml_path))[0]
        parts = stem.rsplit('_', 2)
        if len(parts) == 3:
            return parts[0], parts[1], parts[2]
        return stem, '', ''

    def _get_map_file_candidates(self):
        candidates = []
        source_root = _clean_text(self._source_dir)
        if source_root:
            candidates.append(os.path.join(source_root, self._map_file_name))
            source_parent = os.path.dirname(source_root)
            if source_parent:
                candidates.append(os.path.join(source_parent, self._map_file_name))
        unique_candidates = []
        for candidate in candidates:
            norm_candidate = os.path.normpath(candidate)
            if norm_candidate not in unique_candidates:
                unique_candidates.append(norm_candidate)
        return unique_candidates

    def _infer_category_from_map_key(self, map_key, map_value=None):
        if isinstance(map_value, dict):
            for category_key in ['category', 'family_category', 'discipline', 'group']:
                candidate = _clean_text(map_value.get(category_key))
                if candidate:
                    return candidate

        key_text = _clean_text(map_key)
        if not key_text:
            return ''

        normalized_key = key_text.replace('/', os.sep).replace('\\', os.sep)
        key_dir = os.path.dirname(normalized_key)
        if key_dir and key_dir not in ['.', os.sep]:
            top_dir = _clean_text(key_dir.split(os.sep)[0])
            if top_dir:
                return top_dir

        stem = os.path.splitext(os.path.basename(normalized_key))[0]
        if not stem:
            return ''

        token = _clean_text(stem.split('_', 1)[0])
        if not token:
            token = _clean_text(stem.split('-', 1)[0])
        if not token:
            return ''

        # Keep broader family "kind" for values like EF-U -> EF.
        return _clean_text(token.split('-', 1)[0]) or token

    def _infer_category_from_filename(self, file_name):
        name = _clean_text(file_name)
        if not name:
            return ''

        stem = os.path.splitext(os.path.basename(name))[0]
        if not stem:
            return ''

        token = _clean_text(stem.split('_', 1)[0])
        if not token:
            token = _clean_text(stem.split('-', 1)[0])
        if not token:
            return ''

        # Keep broader family "kind" for values like EF-U -> EF.
        return _clean_text(token.split('-', 1)[0]) or token

    def _load_map_categories(self):
        if self._map_categories_by_file is not None:
            return self._map_categories_by_file

        categories_by_file = {}
        for map_path in self._get_map_file_candidates():
            if not os.path.isfile(map_path):
                continue
            map_payload = self._load_yaml_dict(map_path)
            if not isinstance(map_payload, dict):
                continue

            for map_key, map_value in map_payload.items():
                key_name = os.path.basename(_clean_text(map_key))
                if not key_name:
                    continue
                if not key_name.lower().endswith('.yaml'):
                    continue
                map_category = self._infer_category_from_map_key(map_key, map_value)
                if not map_category:
                    continue
                categories_by_file[key_name.lower()] = map_category

            if categories_by_file:
                break

        self._map_categories_by_file = categories_by_file
        return categories_by_file

    def _load_yaml_dict(self, yaml_path):
        try:
            loaded = yaml.load_as_dict(yaml_path)
        except Exception:
            loaded = None
        if isinstance(loaded, dict):
            return loaded
        return {}

    def _read_first_nonempty(self, payload, keys):
        if not isinstance(payload, dict):
            return ''
        for key in keys:
            value = _clean_text(payload.get(key))
            if value:
                return value
        return ''

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

        self._map_categories_by_file = None
        map_categories = self._load_map_categories()
        entries = []
        for walk_root, _, files in os.walk(self._source_dir):
            for file_name in files:
                file_lower = file_name.lower()
                if not file_lower.endswith('.yaml'):
                    continue
                if file_lower == self._map_file_name:
                    continue

                yaml_path = os.path.join(walk_root, file_name)
                loaded_yaml = self._load_yaml_dict(yaml_path)
                family_token, client_name, project_number = self._parse_filename_tokens(yaml_path)
                family_name = self._read_first_nonempty(loaded_yaml, ['family', 'family_name'])
                if not family_name:
                    family_name = family_token.replace('_', ' ').strip() or family_token

                yaml_client = self._read_first_nonempty(
                    loaded_yaml,
                    ['client', 'client_name', 'project_client', 'project_client_name']
                )
                yaml_project_number = self._read_first_nonempty(
                    loaded_yaml,
                    ['project_number', 'project', 'project_no', 'project_num']
                )
                if yaml_client:
                    client_name = yaml_client
                if yaml_project_number:
                    project_number = yaml_project_number

                map_category = map_categories.get(file_name.lower(), '')
                inferred_category = self._infer_category_from_filename(file_name)
                resolved_category = map_category or inferred_category or self._category_for_path(yaml_path)

                entries.append(
                    FamilyDatabaseEntry(
                        key=yaml_path,
                        family_name=family_name,
                        client_name=client_name,
                        project_number=project_number,
                        category=resolved_category,
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
        loaded = self._load_yaml_dict(entry.key)
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
