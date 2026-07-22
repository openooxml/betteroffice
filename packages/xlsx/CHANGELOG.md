# @betteroffice/xlsx

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
- 793b761: Pending agent proposals render as in-cell tracked-change ghosts painted by the engine: the new value in green above the old value struck through in red, repainting immediately on propose, accept, and reject. Display-list text commands now serialize camelCase so cell fonts, sizes, and strike/underline offsets reach the canvas, and uninstalled workbook fonts fall back to sans-serif instead of the browser serif default.

## 0.0.6

### Patch Changes

- a34e721: Add deterministic Yrs replicas, bounded and validated sync-v1 exchange, a
  transport-agnostic npm collaboration provider, and React peer-update repainting.
  Collaborative sessions support nonstructural cell and dimension edits; inverse-op
  undo and redo remain disabled until a Yrs-aware undo manager can preserve
  concurrent edits.

## 0.0.5

### Patch Changes

- e8678aa: Route workbook sessions through the shared Rust facade and harden editing, calculation, and rendering limits.

## 0.0.4

### Patch Changes

- 6a1ab98: Publish the spreadsheet packages as ESM-only and load the WebAssembly core as a separate asset.

## 0.0.3

### Patch Changes

- 68d15b8: Fix `@betteroffice/xlsx-react` so its dependency on `@betteroffice/xlsx` resolves to the matching published version.
