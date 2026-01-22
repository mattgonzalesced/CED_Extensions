## Goal
Add a new "Place Linked Elements" workflow that lets users select specific host
profile names (from the active YAML) via checkboxes, so only those profiles are
placed instead of always placing every profile in storage.

## Plan
1) Locate the current Place Linked Elements entry point and confirm where
   profile names are sourced (active YAML -> equipment definitions).
2) Define the new button entry point ("Place Linked Elements (Profile Filter)")
   and a checkbox selection dialog listing truth-source groups, defaulting to
   none checked, with search and select-all/none actions.
3) Map selected truth-source groups back to the concrete CAD/profile names
   used by placement so only those profiles generate rows.
4) Filter the placement inputs (equipment_names, selection_map, rows) to the
   chosen profile set and keep summaries scoped to selected profiles only.
5) Validate the workflow messaging: no YAML mutations, clear output on skipped
   profiles, and graceful handling when no matches exist.

## Risks
- Profile naming and truth-source grouping may hide duplicates; filtering by
  display name could place unexpected profiles or skip merged/child entries.
- Large profile sets could make the checkbox UI sluggish or hard to use without
  search and bulk actions.
- Users may expect category-filter and standard Place Linked Elements to behave
  consistently; divergent behavior could cause confusion.

## Open Questions
- None
