# @betteroffice/xlsx-react

## 0.0.7

### Patch Changes

- 793b761: Render pending proposals as Word-style tracked changes: struck-through old
  values with a red run highlight, new values in green with a dashed underline
  and green run highlight, laid out side by side or new-over-old and following
  cell alignment. Proposal staging recalculates the formula graph and ghosts
  downstream dependents whose computed values change, proposal edits can carry
  a number format, and no-op proposals render unmarked.
- c6ad184: Add a Google Sheets-style toolbar to the XLSX editor backed by new engine
  APIs for range styling, number formats, selection-format aggregation, format
  painting, merge queries, and history state. Formatting is fully collaborative
  through a content-addressed style catalog (collaboration schema v3; v2 state
  does not migrate). Merging replaces intersecting ranges like Excel, parsing
  repairs overlapping merges in third-party files, and display-list font fields
  now serialize correctly so styled text renders with its real font, size, and
  weight.
- Updated dependencies [793b761]
- Updated dependencies [c6ad184]
- Updated dependencies [793b761]
  - @betteroffice/xlsx@0.0.7
  - @betteroffice/xlsx-i18n@0.0.7

## 0.0.6

### Patch Changes

- a34e721: Add deterministic Yrs replicas, bounded and validated sync-v1 exchange, a
  transport-agnostic npm collaboration provider, and React peer-update repainting.
  Collaborative sessions support nonstructural cell and dimension edits; inverse-op
  undo and redo remain disabled until a Yrs-aware undo manager can preserve
  concurrent edits.
- 69d62f1: Refine the XLSX and PPTX editor toolbars with compact DOCX-style control rails,
  grouped icon actions, and responsive value fields.
- Updated dependencies [a34e721]
  - @betteroffice/xlsx@0.0.6
  - @betteroffice/xlsx-i18n@0.0.6

## 0.0.5

### Patch Changes

- Updated dependencies [e8678aa]
  - @betteroffice/xlsx@0.0.5

## 0.0.4

### Patch Changes

- 6a1ab98: Publish the spreadsheet packages as ESM-only and load the WebAssembly core as a separate asset.
- Updated dependencies [6a1ab98]
  - @betteroffice/xlsx@0.0.4

## 0.0.3

### Patch Changes

- 68d15b8: Fix `@betteroffice/xlsx-react` so its dependency on `@betteroffice/xlsx` resolves to the matching published version.
- Updated dependencies [68d15b8]
  - @betteroffice/xlsx@0.0.3
