# Vendored third-party libraries

Each subdirectory is a vendored copy of a Python package, included
because pyRevit's bundled CPython 3 engine has no exposed `site-packages`
and there is no other guaranteed way to make these libraries available
to scripts in the panel.

## yaml/

PyYAML pure-Python source (no C extension), copied from
`PyYAML 6.0.3` (PyPI). License: MIT.

The C extension `_yaml.*.pyd` was intentionally omitted; PyYAML falls
back to pure-Python automatically when `_yaml` cannot be imported.

To upgrade: replace the entire `yaml/` subdirectory with the contents
of a fresh PyPI wheel's `yaml/` package, again skipping any binary
`.pyd` / `.so` and `__pycache__` directories.
