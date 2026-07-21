# @betteroffice/xlsx

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
