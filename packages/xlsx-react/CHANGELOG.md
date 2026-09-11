# @betteroffice/xlsx-react

## 0.1.0

### Minor Changes

- fe1145a: A chart on a worksheet is now a selectable object instead of a picture the click falls through. Every frame publishes each chart's id, rect, clipped hit area and whether it can be repinned, and chrome resolves a press against that frame's own regions, so the answer cannot drift from the pixels. Clicking a chart selects it and outlines it; dragging it, or nudging it with the arrow keys, slides it through the op log as one undoable edit; and the moved anchor — cell and EMU offset alike — is written back into the drawing part on save, synthesising the `colOff`/`rowOff` a drawing omitted rather than saving a move that lost half of itself. Clicking off the chart restores the cell selection, and while a chart is selected the keyboard no longer reaches the cells hidden behind it.

  A chart the renderer could not draw is selectable too: it degrades to a placeholder but still occupies its space, so it stays an object that can be picked up and moved out of the way.

  Two limits worth stating. A chart pinned by an absolute anchor can be selected but not moved: its position lives in attributes the writer cannot rewrite, and the frame reports this as `movable: false` so the UI never offers the drag. And moving a chart is a standalone-session edit — chart state syncs as one blob per sheet, so a collaborative session refuses it exactly as it refuses freeze panes and hyperlinks.

  Minor rather than patch: `DisplayList.charts` changes its element contract. The elements gain `id`, `rect`, `clip` and `movable` beside `placeholder`, and lose `Eq` on the Rust side, so anything that _constructs_ one must be updated — `ChartA11yAttrs` survives only as an alias for readers. `Op::SetChartAnchor` is a new variant and breaks exhaustive matches on `Op`.

### Patch Changes

- 148ba02: A workbook snapshot persisted by the previous release opens again. A replica bootstraps the workbook it opened into a document whose client ID is the head of the base fingerprint, and that fingerprint hashes the collaboration schema version — so raising the version for charts gave every workbook a different bootstrap identity. A snapshot an earlier release wrote no longer deduplicated against the one this build seeds: the two bases doubled up, one was tombstoned by client ID, and restoring reported that the shared workbook structure had changed or silently handed back the pristine file.

  A replica that has not been edited yet now takes a whole snapshot as its state and upgrades it in place, rather than merging it against a bootstrap it can never agree with. Where the two bootstraps do agree the snapshot and the merge describe the same document, so this is never the worse answer. The upgraded state, not the snapshot, is what peers are told about: the upgrade writes new structs, and an incremental update that later builds on them would sit unintegrated forever on a peer that never received them. What the frozen structure describes — sheet order and names, merges, freeze panes, hyperlinks and charts — must still match, disregarding the shared-type identities a replaced bootstrap changes by construction. A snapshot that fails that, or a whole document this build cannot read, is now an error rather than a silent no-op.

  A charted workbook pairs with a pre-chart snapshot again. Such a snapshot carries no chart state to disagree about — charts come from the file the replica opened, keyed to the sheet they were parsed from — so refusing the pairing only made charted workbooks the ones that could never be restored.

  Workbooks with a hidden row or column restore too. This release began modelling those as a zero dimension where earlier ones recorded nothing at all, and both dimension maps are fingerprinted, so such a workbook could not be recognised as the base its own snapshot had started from. The dimensions an earlier release would have stored are now read from the source sheet alongside the current ones and accepted as a legacy fingerprint, and restoring puts the hidden dimensions back rather than letting the row silently unhide.

  Each feature is now pinned to the schema that introduced it rather than to whichever schema is current, so the next version bump cannot reclassify the newest schema as predating charts and manufacture this same failure again.

- Updated dependencies [692f2c7]
- Updated dependencies [ab39d50]
- Updated dependencies [56fde13]
- Updated dependencies [0f1ae6b]
- Updated dependencies [fe1145a]
- Updated dependencies [d143a82]
- Updated dependencies [5989039]
- Updated dependencies [64eecea]
- Updated dependencies [a7b8062]
- Updated dependencies [623ee21]
- Updated dependencies [36c87cf]
- Updated dependencies [f5d1b03]
- Updated dependencies [9534169]
- Updated dependencies [148ba02]
  - @betteroffice/xlsx@0.1.0
  - @betteroffice/xlsx-i18n@0.1.0

## 0.0.8

### Patch Changes

- 4e04087: Formulas referencing defined names now resolve correctly, frozen panes render, and hyperlinks survive the round trip. The collaboration schema advances to version 5 and upgrades version 3 and 4 snapshots when read, so a client on this release cannot share a collaboration room with an older one: upgrade every peer together.
- 0d3baa1: Collaborative presence: remote collaborators' cell and range selections render as colored outlines with name flags, plus toolbar avatar chips; worksheets expose stable collaborative ids so presence survives sheet renames.
- Updated dependencies [4e04087]
- Updated dependencies [47c37b0]
- Updated dependencies [0d3baa1]
  - @betteroffice/xlsx@0.0.8
  - @betteroffice/xlsx-i18n@0.0.8

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
