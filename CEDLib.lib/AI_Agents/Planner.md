## Goal
Update Update Profiles so keynotes and text notes are captured by proximity:
if a keynote or text note is within 5 feet of an element with Element_Linker,
attach it to that element’s profile, choosing the closest host and preventing
duplicates across profiles by location.

## Plan
1) Identify where Update Profiles currently gathers keynotes and text notes and
   how it maps them to profiles (hosted vs non-hosted elements).
2) Define proximity capture rules: 5 ft radius in XY only (ignore Z), choose
   closest host element, and a location-based de-duplication key so a note is
   stored in only one profile.
3) Build or reuse a host-element spatial lookup (Element_Linker only) and
   update note collection to resolve each note to the nearest eligible host
   using the note insertion point for XY distance.
4) Enforce uniqueness by location across profiles (track assigned XYZ keys
   during the scan) and skip notes already assigned.
5) Handle exact distance ties by prompting the user to select the host profile
   for the keynote/text note.
6) Ensure GA_Keynote Symbol_CED keynotes are stored without adding a 90-degree
   rotation (no extra rotation applied during capture).
7) Verify placement updates: keynotes/text notes near multiple hosts resolve to
   the closest (or user-selected on ties), and notes outside 5 ft are ignored.

## Risks
- Ambiguous distances or floating-point noise can cause inconsistent host
  selection or duplicate detection unless XY tolerances are defined.
- Large models could make proximity searches slow without spatial indexing.
- Notes near linked model boundaries might match unintended hosts if level or
  view context is ignored.

## Open Questions
- None (no existing profile notes to migrate).
