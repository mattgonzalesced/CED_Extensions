# PR Summary

## Overview
This update includes reliability fixes in Move/Editor/Circuit Browser behavior and a full rewrite of release metadata automation so PR metadata updates are managed by one replaceable bot commit.

## Application Fixes

### 1. Move Selected Circuits crash fix
File:
- `CEDLib.lib/CEDElectrical/Application/services/move_circuits_to_panel_service.py`

Change:
- Added missing initialization:
  - `option_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)`

Impact:
- Fixes runtime error:
  - `global name 'option_filter' is not defined`
- Covers both entry points that share the same operation/service path:
  - Ribbon Move Selected Circuits
  - Circuit Manager Move Selected Circuits

### 2. Edit Circuit Properties neutral/IG toggle UI fix
File:
- `CEDLib.lib/CEDElectrical/ui/circuit_properties_editor.py`

Change:
- In `_set_combo_value`, when preview value is `"-"` and the combo has no `"-"` option, clear the selection/text instead of preserving stale value.

Impact:
- Neutral/IG wire-size dropdown no longer shows stale size after unchecking include toggle.
- UI now reflects excluded state immediately.

### 3. Circuit Browser action hardening
File:
- `CED ElecTools.extension/AE pyTools.tab/Electrical.panel/Circuit Manager.pushbutton/CircuitBrowserPanel.py`

Changes:
- `_raise_action_operation` now checks `raise_operation(...)` result and surfaces queue failure instead of always returning success.
- Added stale-item pruning (`_prune_stale_items`) before opening action windows for:
  - Mark Existing
  - Add/Remove Neutral
  - Add/Remove Isolated Ground
  - Breaker autosize
  - Edit Circuit Properties
- Added try/catch guard + alert around opening Neutral and IG action windows.

Impact:
- Prevents silent queue failures.
- Reduces invalid/deleted-element errors.
- Improves operator feedback on action startup failures.

## CI / Release Metadata Automation

### 4. Workflow rewrite and consolidation
File:
- `.github/workflows/bump-version.yml`

Replaced prior workflow with a single consolidated metadata workflow that:
- Triggers on PR events:
  - `opened`, `synchronize`, `labeled`, `unlabeled`
- Targets branches:
  - `develop`, `main`
- Skips bot actor:
  - `github.actor == 'github-actions[bot]'` is blocked
- Uses `github.base_ref` to branch behavior.

Branch behavior:
- `develop`:
  - updates `build` only
  - does not modify `toolbar_version`
- `main`:
  - requires exactly one release label:
    - `release: major`
    - `release: minor`
    - `release: patch`
  - reads base version from `origin/main` copy of `about.yaml`
  - computes next semver from base:
    - major -> `X+1.0.0`
    - minor -> `X.Y+1.0`
    - patch -> `X.Y.Z+1`
  - updates both `toolbar_version` and `build`

Single replaceable bot commit behavior:
- Commit message is fixed:
  - `Bot: update release metadata`
- Detects prior matching bot commit(s) in PR branch history (relative to base branch).
- Drops prior bot metadata commit(s) before recomputing metadata.
- Creates at most one fresh bot metadata commit at tip.
- Pushes with:
  - `git push --force-with-lease`

Logging:
- Added `::notice::` / `::warning::` logs for:
  - target branch
  - labels
  - base version
  - bot commit found/removed
  - computed version
  - build timestamp
  - commit/push skipped vs performed

### 5. New helper script for deterministic file updates
File:
- `.github/scripts/update_release_metadata.py` (new)

Behavior:
- Updates/inserts `build: YYYYMMDDHHmm`
- Optionally updates/inserts `toolbar_version: X.Y.Z`
- Preserves existing file layout/order semantics without relying on external action scripts.

Target metadata file:
- `AE pyTools.extension/AE pyTools.Tab/CED Tools.panel/About.pushbutton/about.yaml`

## Validation Performed
- YAML parse check for workflow (`yaml.safe_load`).
- Python syntax checks:
  - `py_compile` on helper script
  - `py_compile` on modified Python app files
- Local helper script smoke test against sample `about.yaml` content.

