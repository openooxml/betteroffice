# @betteroffice/xlsx-react

## 0.3.0

### Minor Changes

- 8132fe0: Add version-checked, all-or-nothing XLSX edit batches. `WorkbookHandle.version()`, `readCells()` and `findText()` return values, formulas and display text with the session version they were read at; `validateEdits()` and `applyEdits()` resolve every step against that version and either commit cell inputs, formulas, number formats and style patches as one recalculated change and one undo step, or return a typed refusal (`stale-version`, `missing-target`, `content-mismatch`, `overlapping-steps`, `locked-target`, `unsupported`, `invalid-step`, `limit-exceeded`) with the workbook untouched. Steps can guard each cell's value, formula or display text, the experimental `history: "none"` keeps a batch out of undo, and volatile functions see only the request's `nowSerial`. `XlsxEditorApi` gains the same operations, committing pending drafts, nudges and pastes first, rejecting with the newly exported `XlsxCommandAdmissionError` (`input-failed`, `gesture-active`, `document-replaced`) when that input cannot be written, a chart drag is unfinished or the workbook was replaced, and refusing with `read-only` while the editor is read-only. The Rust `Workbook` and the Python bindings expose the same API. Update observers now run after recalculation, cell batches, raw operations and proposal acceptance stage the authority change before adopting it, and `StaleProposalError` gains `targets` naming each drifted cell's sheet. Edits to cells that carry formats created in the same session now reach the shared Yrs state with the right format, and a value-changing `recalculate_all` notifies Rust observers with a new `UpdateOrigin::Recalculation` event that carries no update. The editor's cut now clears only the cells selected when it was accepted, and only while that workbook is still open and writable. The Rust `EditOperation` and `EditFailureCode` enums are non-exhaustive.
- f07e295: Expose the XLSX editor's commands as `XlsxEditorApi.commands`: serializable state with a stated reason for every disabled command, descriptors with shortcuts, and `execute`, which runs in order with the pastes, cell entries and chart moves accepted before it, ends an IME composition and writes the text typed so far, and checks availability again, failing with `target-changed` when its selection moved meanwhile; print waits for the canvas to paint that text and fails with `render-failed` otherwise. The synchronous `api.save()` throws the new `XlsxSaveRefusedError` (`input-pending`, `input-failed`) instead of returning bytes without accepted input. A cell entry the engine refuses keeps the cell editor open with its error, and Save and Mod+S fail with `input-failed` until the entry is corrected or cancelled with Escape. Hosts compose built-in controls with their own actions through the `toolbar` and `showToolbar` props, `XlsxCommandProvider`, `useXlsxCommands`/`useXlsxCommandState`/`useXlsxCommand`, `EditorToolbar mode="commands"` (with `EditorToolbar.FormulaBar`), `ToolbarCommand`, `ToolbarCommandButton`, `ToolbarCommandSelect` and `ToolbarOverflow`, inside or outside the editor. The prop-based `EditorToolbar`, `Toolbar` and `useEditorToolbar` keep working in the default legacy mode. Toolbars, shortcuts and the proposals panel share one gate: read-only mode, cell selection, merge shape, collaboration and PNG support each disable commands with a coded reason. Shortcuts dispatch from the descriptors with platform-aware labels (now including Mod+B and Mod+I), only for the editor that owns the event, and leave text undo to the cell editor. Narrow toolbars move trailing groups into a keyboard-accessible More menu with radio choices, custom size and zoom prompts and color pickers, and `ToolbarDropdown` popups follow menu or dialog semantics by their content. New locale keys cover command labels and disabled reasons. The command and toolbar composition API is experimental and may change in minor releases.
- 5206ccf: Add host-controlled viewing mode, change notifications, imperative save, and cell selection APIs.
- 667fae3: Host editor plugins in `XlsxEditor` through the new `plugins`, `pluginGrants` and `onPluginError` props. The plugin API is experimental and may change in minor releases. A plugin created with `defineXlsxPlugin` can contribute a docked panel (left, right or bottom of the grid, above the sheet tabs), an overlay on the grid, and commands registered as `plugin:<pluginId>/<id>`, whose results carry the plugin's own failure codes or a refused edit batch unchanged, with toolbar entries (`XlsxPluginToolbar` places them in replacement chrome) and shortcuts; `ToolbarCommandButton`, `ToolbarCommand`, `useXlsxCommand` and `useXlsxCommandState` accept those ids. Each activation receives `load`, `document-change` (the committed version after recalculation), `selection-change` (sheet, cells in their direction, chart), `mode-change`, `layout-change` and `grants-change` events, restricted clients (versioned `readCells`, `findText` and `validateEdits`, commands, granted edit batches, and `selectCells` and `scrollToCell` navigation that keeps focus unless asked), grid geometry that returns cell and range rectangles in overlay pixels for the painted frame, zoom, scroll and frozen panes, version-guarded state and `onCleanup` disposers. Plugins read, validate and navigate by default; built-in commands and edit batches need an explicit grant, rechecked with `readOnly` immediately before every write, and mutating built-in commands refuse plugins with `unsupported-policy`. Every contribution renders behind its own error boundary with a plugin-scoped command context and without the editor's formula bar binding, its keys and clicks never reach the grid, and a failing plugin is stopped and reported without affecting the editor or other plugins. Plugin shortcuts must use Mod or Alt, or a function key; built-in shortcuts win and a clashing plugin shortcut, including one the grid handles itself, is reported and not bound. The proposals panel now opens inside the grid area. New locale keys cover plugin panels, the plugin toolbar group and the new disabled reasons.

### Patch Changes

- Updated dependencies [6c30f4a]
- Updated dependencies [9639a6b]
- Updated dependencies [80341ac]
- Updated dependencies [e7f4868]
- Updated dependencies [d1c6a55]
- Updated dependencies [3153c95]
- Updated dependencies [ae8c931]
- Updated dependencies [6bd3783]
- Updated dependencies [c01cc85]
- Updated dependencies [0f23444]
- Updated dependencies [21d83b3]
- Updated dependencies [21d83b3]
- Updated dependencies [91554bf]
- Updated dependencies [d1c6a55]
- Updated dependencies [8132fe0]
- Updated dependencies [f07e295]
- Updated dependencies [d1c6a55]
- Updated dependencies [d1c6a55]
- Updated dependencies [ddbfd29]
- Updated dependencies [b56efb3]
- Updated dependencies [b4db492]
- Updated dependencies [65f0ce9]
- Updated dependencies [667fae3]
- Updated dependencies [6d28c9d]
- Updated dependencies [d1c6a55]
- Updated dependencies [dfb5c82]
- Updated dependencies [d1c6a55]
- Updated dependencies [d1c6a55]
- Updated dependencies [9a6cdd5]
- Updated dependencies [adb8999]
- Updated dependencies [d1c6a55]
- Updated dependencies [6d28c9d]
- Updated dependencies [6c30f4a]
- Updated dependencies [bcc90ba]
- Updated dependencies [35f9cec]
- Updated dependencies [d1c6a55]
- Updated dependencies [c01cc85]
- Updated dependencies [d1c6a55]
- Updated dependencies [b56efb3]
- Updated dependencies [b56efb3]
  - @betteroffice/xlsx@0.3.0
  - @betteroffice/xlsx-i18n@0.3.0

## 0.2.1

### Patch Changes

- 16d33a7: Fix locale declarations for TypeScript consumers with `skipLibCheck: false` and update React editors to depend on the corrected i18n packages.
- Updated dependencies [16d33a7]
  - @betteroffice/xlsx-i18n@0.2.1
  - @betteroffice/xlsx@0.2.1

## 0.2.0

### Patch Changes

- 5069ad2: Keep accepted spreadsheet proposals undoable in collaborative sessions, preserve pending proposals through remote edits, and require a refreshed review when calculated previews change. Reject document suggestions that overlap partially tracked text. Existing public signatures and wire fields remain unchanged.
- 93971b5: Remove outdated early-release warnings from package READMEs and link the JavaScript guide and changelogs.
- Updated dependencies [5069ad2]
- Updated dependencies [93971b5]
- Updated dependencies [fb06233]
- Updated dependencies [d926fb0]
- Updated dependencies [cae162d]
- Updated dependencies [8a27bb7]
- Updated dependencies [7f062e0]
- Updated dependencies [0c5c4fc]
- Updated dependencies [5798031]
- Updated dependencies [13016f2]
- Updated dependencies [e8c4f5b]
  - @betteroffice/xlsx@0.2.0
  - @betteroffice/xlsx-i18n@0.2.0

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
