# CED pyTools 2.0 Release Description
## Baseline and Scope
This release description is based on `main` -> `release/2.0-cleanup`.

- Commits ahead of `main`: **118**
- Merge commits in scope: **11**
- Net file delta: **779 files changed** (`55,073` insertions, `221,045` deletions)
- Major merged streams represented in history:
  - `MergeCircuits`
  - `RenameCircuitByDeviceParameter`
  - `Refrigeration-Tools`
  - `Super-Circuit`
  - multiple `develop` and `main` sync merges into the release line

## What 2.0 Changes at a High Level
Version 2.0 is a structural and workflow release, not a minor patch. It delivers:

- Trade-segmented tool distribution (`CED ElecTools.extension` and `CED MechTools.extension`)
- A modernized electrical circuit management stack (browser, alerts, staged operations, finder workflows)
- Expanded mechanical/refrigeration operations toolkit
- New AE pyTools utility layer improvements (Quick Dimension, view operations, panel reorg)
- Startup/telemetry infrastructure hardening for usage capture and transfer

## Trade-Segmented Extensions (Electrical + Mechanical Split)
Electrical and mechanical toolsets are now separated into dedicated extension roots:

- `CED ElecTools.extension`
- `CED MechTools.extension`

This separation allows teams to enable/disable by trade in pyRevit instead of carrying one monolithic mixed toolbar. The result is cleaner role-based deployment and lower UI clutter for each discipline.

## Electrical: New and Upgraded Features
## Electrical managers and browser workflows
- **Circuit Manager** (dockable pane): promoted to a first-class manager surface with filtering, staged actions, selection utilities, and operation orchestration.
- **Alerts Manager**: dedicated alert browsing/review workflow with operation-backed refresh/recalculate paths.
- New operation flow built around shared external-event execution and common request/runner handling.

## Circuit operations and editing workflows
- Expanded/updated circuit workflows including:
  - Add Spares and Spaces
  - Batch Swap Circuits
  - Edit Circuit Properties
  - Move Selected Circuits
  - Merge Circuits
  - Rename Circuits by Device Parameter
  - Create Dedicated Circuits
  - Load Electrical Parameters
  - Calculate Circuits updates with settings integration
- Reliability fixes for staged include/remove logic (Neutral/IG), recalculation paths, and panel interactions.

## Circuit/device finding and model navigation
- **Circuit Element Finder / Circuited Device Finder** added for direct model visibility workflows from selected circuits and browser actions.
- **Find Circuited Elements** and related selection actions were aligned with the manager workflow for faster downstream/upstream navigation.

## Wire + one-line + QC support
- Wire tool path improvements, including updated behavior around native wire generation flow and config handling.
- **Sync One-Line Data** retained and integrated into the reorganized electrical stack.
- QC tools retained under the electrical extension:
  - Color Circuits Check
  - Color Circuits by Panel
  - Electrical System Check
  - Panel Report

## Electrical UX/platform upgrades
- Theme/resource handling consolidated across circuit and alert windows.
- Shared UI resources and style system reduced drift between tools.
- Path and panel-level command organization updated (managers promoted to top-level).

## Mechanical/Refrigeration: New and Upgraded Features
Mechanical tools are now centralized under `CED MechTools.extension` with both piping utilities and ref operations:

- **Set Pipe System**
- **ConnectTo**
- **DisConnect**
- **Element 3D Rotation**
- **MakeParallel**
- **Transition**
- **Name Piping Systems**
- **Place all Coils**
- **Space Coils**
- **System Tagger**
- **Print Pipe Data**

Behavioral improvements reflected in history include:

- Name Piping Systems iteration and sequencing improvements
- Space Coils distribution updates and wall/space boundary behavior refinements
- Print Pipe Data introduced to support downstream piping-system QA workflows

## AE pyTools (Core) Additions and Reorganization
In parallel to the trade split, AE pyTools received direct utility upgrades and panel cleanup:

- **QuickDimension** added (`MiscTools1.stack`)
- **Views.pulldown** enhanced/reorganized with:
  - Copy VG Settings to View Templates
  - Toggle Grid Bubbles
  - Unhide All in Active View
  - XBG Grey All Layers
- **Copy Import Visibility** added under MEP Automation misc operations
- **About** button/window added to CED Tools
- Miscellaneous/tool stacks reorganized for clearer navigation and reduced legacy clutter

## Telemetry System (Startup + Close Transfer Pipeline)
2.0 formalizes telemetry handling in startup and shutdown flow:

- Startup ensures telemetry output directory exists at:
  - `%APPDATA%\pyRevit\Extensions\CED_pyTelemetry`
- Startup sets pyRevit telemetry configuration (active/UTC/hooks/file-dir, with app telemetry disabled by default).
- On Revit close, telemetry files are moved to ACC usage storage when available:
  - `...\CED Content Collection\Project Files\03 Automations\Usage\<username>`
- If ACC is not synced/available, the system safely skips transfer and retains local telemetry files.
- ACC location checks include both local ACC path patterns used by the team environment.

This gives a dependable local-first capture path with deferred cloud transfer when project sync is available.

## Architecture and Foundation Upgrades
The release introduces broader internal platform work to support stability and future scale:

- CEDElectrical application layering (contracts, DTOs, operations, services, infrastructure)
- ExternalEvent gateway standardization for modeless UI operations
- Repository/writer abstractions for circuit and panel workflows
- Shared UI resource framework and theme bridge support
- Large cleanup/removal of legacy/prototype/pydev-era artifacts

## Breaking/Behavioral Notes for 2.0
- Command locations changed due to manager promotion and panel restructuring.
- Electrical tools are no longer hosted in the old AE pyTools electrical location; they now live under `CED ElecTools.extension`.
- Mechanical/refrigeration workflows are now isolated in `CED MechTools.extension`.
- Startup now owns dockable-pane registration patterns for core panes, reducing runtime registration drift.

## Release Positioning
CED pyTools 2.0 is a full platform refresh with discipline-based extension boundaries, an upgraded electrical operations stack, expanded mechanical tooling, and hardened startup/telemetry behavior. The release prioritizes maintainability, role-based deployment, and high-throughput production workflows over backward-compatible command layout stability.
