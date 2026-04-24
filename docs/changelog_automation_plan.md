# Changelog Automation Plan (On Hold)

## Status
- State: On hold
- Decision date: 2026-04-24
- Active now: release/main PR labeling + manual version/build bump flow
- Deferred: automatic changelog generation

## Goal (Deferred)
- Automatically build release notes from merged `develop` work.
- Keep release notes consistent, concise, and traceable to PRs.
- Group notes by:
  - Highlights (top section)
  - Tool (section per tool)
  - Change entries under each tool in format:
    - `<dev tag>` - Description ([#PR](link))

## Proposed Label Model (Deferred)
- Required exactly one on PRs into `develop`:
  - `type: bug fix`
  - `type: new feature`
  - `type: enhancement`
  - `type: deprecated`
- Optional:
  - `highlight`

## Proposed Data Model (Deferred)
- Keep one durable source file (no one-file-per-PR sprawl):
  - `docs/changelog_index.json` (or `.jsonl`)
- Each record keyed by PR number, with:
  - PR number, title, URL
  - merged timestamp
  - type label
  - highlight flag
  - short description (from PR body or curated field)
  - affected files
  - inferred tool(s) from file path map (or explicit tool labels)

## Proposed Rendered Outputs (Deferred)
- `docs/changelog.md` for human browsing.
- `release_notes.md` generated at release time from changelog index since last tag.

## Proposed Workflow Integration (Deferred)
1. On PR merge into `develop`, upsert one changelog record for that PR.
2. Regenerate `docs/changelog.md` from index.
3. Commit updated changelog artifacts.
4. During release creation, generate release notes from records between last tag and current tag.

## Open Decisions
- Source of truth for description: PR body vs structured template section.
- Tool mapping strategy: path rules vs explicit `tool:*` labels.
- Release note window: tag-to-tag vs date window.
- Edit model: fully automated vs maintainer review gate before publish.

## Implementation Backlog (When Resumed)
1. Create `.github/scripts/changelog.py` with commands:
   - `record-pr`
   - `render`
2. Define label and PR template policy for `develop` PRs.
3. Add/adjust workflow step(s) to call `record-pr` on merged PRs to `develop`.
4. Add release workflow step to render `release_notes.md` from changelog index.
5. Add tests for parser/render logic and idempotent re-runs.
