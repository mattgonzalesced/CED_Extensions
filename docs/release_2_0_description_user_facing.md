# CED pyTools 2.0 Release Description (User-Facing)
## Baseline and Scope
This summary focuses on what users will notice in day-to-day work: new tools, modified workflows, and where to find them.

## What 2.0 Changes at a High Level
CED pyTools 2.0 is a major workflow release with stronger electrical management, expanded mechanical operations, and cleaner by-trade deployment.

Tag format used below:

- **[NEW]** = brand-new tool or workflow
- **[MODIFIED]** = existing tool with meaningful behavior/UI changes
- **[BETA]** = available for production testing while still actively iterated

## Trade-Segmented Extensions (Electrical + Mechanical Split)
Electrical and mechanical tools are now split into dedicated extensions:

- `CED ElecTools.extension`
- `CED MechTools.extension`

User impact:

- Teams can enable/disable by trade.
- Ribbons are cleaner for each discipline.
- Rollout is simpler for role-specific users.

## Electrical: New and Upgraded Features
## Electrical managers and browser workflows
- **[NEW]** **Circuit Manager**: centralized dockable browser for filtering, selecting, reviewing, and running circuit actions.
- **[NEW]** **Alerts Manager**: dedicated alert review window with refresh/recalculate and model selection workflows.

## Circuit operations and editing workflows
- **[MODIFIED]** **Add Spares and Spaces**: new bulk-action UI, faster quick-apply behavior, and improved switchboard handling defaults.
- **[MODIFIED]** **Batch Swap Circuits**: redesigned staged drag/drop workflow across panels and schedules.
- **[MODIFIED]** **Edit Circuit Properties**: cleaner editing experience with faster review/apply cycles.
- **[MODIFIED]** **Move Selected Circuits**: improved handling around spare/space cleanup during moves -> allows circuits to replace default spares/spaces if desired.
- **[NEW]** **Merge Circuits**: move devices from source circuits into a selected target circuit.
- **[BETA]** **Rename Circuits by Device Parameter**: rename load names from connected device data, including template-style naming controls.
- **[MODIFIED]** **Calculate Circuits**: updated settings flow and tighter integration with manager-driven workflows.
- **[MODIFIED]** **Wire Circuited Elements**: Now uses native Revit wire generation algorithm, better configuration options, and stronger fallback resolution when saved wire types are missing.

What changed for users:

- More reliable staged edits before apply.
- More predictable apply/recalculate outcomes.
- Better stability in Neutral/IG modification workflows.

## Circuit/device finding and model navigation
- **[NEW]** **Circuited Device Finder**: quickly show connected devices (and related circuit context) in model views.
- **[MODIFIED]** **Find Circuited Elements**: smoother selection and navigation behavior.

## one-line + QC support
- **[MODIFIED]** **Sync One-Line Data**: retained as a core electrical workflow and aligned with 2.0 circuit operations.
- **[MODIFIED]** **Color Circuits Check**: quickly visualize and validate circuiting conditions by color rules.
- **[MODIFIED]** **Color Circuits by Panel**: color systems by source panel to speed distribution QA.
- **[MODIFIED]** **Electrical System Check**: run circuit/system validation checks to catch coordination issues earlier.
- **[MODIFIED]** **Panel Report**: generate a quick panel-focused review output for schedule and loading checks.

## Electrical UX/platform upgrades
- Circuit and alert windows are more consistent in behavior and layout.
- Manager-first organization makes key actions easier to find.
- UI consistency is improved across related electrical tools.

## Mechanical/Refrigeration: New and Upgraded Features
Mechanical/refrigeration tools are grouped under `CED MechTools.extension`.

### Mechanical Tools
- **[NEW]** **Set Pipe System**: assign/update piping system values across selected elements with less manual cleanup.
- **[NEW]** **ConnectTo**: quickly connect compatible piping elements/connectors in one action flow.
- **[NEW]** **DisConnect**: quickly break connector relationships without manual connector-by-connector edits.
- **[BETA]** **Element 3D Rotation**: rotate selected elements in 3D with faster directional control.
- **[BETA]** **MakeParallel**: align selected runs/elements to parallel orientation with reduced manual adjustment.
- **[BETA]** **Transition**: speed transition placement/alignment workflows for piping connections.

### Refrigeration Tools
**Note:** All refrigeration tools are **[BETA]** and in active testing on select projects.

- **[BETA]** **Name Piping Systems**: apply/refine piping naming logic for refrigeration layouts with improved consistency.
- **[BETA]** **Place all Coils**: automate bulk coil placement workflows with updated placement behavior.
- **[BETA]** **Space Coils**: improve coil spacing distribution and layout consistency.
- **[BETA]** **System Tagger**: streamline system ID tagging workflows for refrigeration deliverables.
- **[BETA]** **Print Pipe Data**: output key piping data for QA review and troubleshooting.

What changed for users:

- More complete piping/refrigeration workflow coverage in one place.
- Better spacing/naming behavior in refrigeration operations.
- Easier QA support with **Print Pipe Data**.

## AE pyTools (Core) Additions and Reorganization
AE pyTools also received direct user-facing updates:

- **[NEW]** **QuickDimension**: bulk dimension generation between multiple elements for quick annotation.
- **[MODIFIED]** **Copy VG Settings to View Templates**: Copy specific overrides from an active view to targeted view templates.
- **[MODIFIED]** **Toggle Grid Bubbles**: quicker batch control of grid bubble visibility.
- **[MODIFIED]** **Unhide All in Active View**: one-click reset for hidden elements in the current view.
- **[MODIFIED]** **XBG Grey All Layers**: streamlined control of imported-layer visibility state.
- **[NEW]** **Copy Import Visibility**: transfer import visibility setup between views.
- **[NEW]** **About**: quick access to release/version information, request/help forms, and learning resources.
- **[MODIFIED]** Miscellaneous stacks were reorganized for easier discovery.

## Telemetry System (Startup + Close Transfer Pipeline)
We introduced a new telemetry system in 2.0 for improved diagnostics and usage metrics.

## Architecture and Foundation Upgrades
2.0 includes broad stability and consistency upgrades under the hood. For users, this translates to fewer edge-case failures and more consistent behavior across tools.

## Breaking/Behavioral Notes for 2.0
- Some commands moved due to panel/manager reorganization.
- Electrical tools are now under `CED ElecTools.extension`.
- Mechanical/refrigeration tools are now under `CED MechTools.extension`.

## Release Positioning
CED pyTools 2.0 is a production-focused release centered on clearer trade workflows, stronger circuit and piping toolsets, and better day-to-day reliability.
