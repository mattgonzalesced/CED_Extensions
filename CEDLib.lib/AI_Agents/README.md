# AI_Agents

This folder is the designated location for:

- **Planning and design documents** for AI-driven or automated workflows (e.g. quality checking, analysis).
- **Agent-related specs** that describe how something should be coded, without containing implementation code.
- **New subfolders and documents** created for CED extensions that fall under the “AI_Agents” scope.

All new folders or documents for this scope should be filed under `CEDLib.lib/AI_Agents`.

## Quality Check Planner role

The **Quality Check Planner** is responsible for:

1. **Planning and documenting** how quality checking and analysis should be coded (architecture, contracts, data flow, rule categories)—without writing implementation code.
2. **Receiving feedback from the MCP code reviewer** and **adjusting coding plans** to reflect their reviews. When the code reviewer provides feedback (e.g. on structure, contracts, testability, or alignment with existing patterns), the planner updates the relevant documents under `CEDLib.lib/AI_Agents` so that coding plans stay accurate and implementable.

Review-driven changes should be applied to the affected planning docs (e.g. `QualityCheck_Coding_Plan.md`) and, when useful, summarized in a **Review log** (see [Review_Log.md](Review_Log.md)).

## Contents

| Document | Purpose |
|----------|---------|
| [QualityCheck_Coding_Plan.md](QualityCheck_Coding_Plan.md) | Plan and documentation for how Revit quality checking and analysis should be coded (MEP, Refrigeration, Fire Protection). |
| [Review_Log.md](Review_Log.md) | Optional log of MCP code reviewer feedback and corresponding updates to coding plans. |
