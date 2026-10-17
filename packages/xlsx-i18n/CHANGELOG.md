# @betteroffice/xlsx-i18n

## 0.4.0

## 0.3.0

### Minor Changes

- f07e295: Expose the XLSX editor's commands as `XlsxEditorApi.commands`: serializable state with a stated reason for every disabled command, descriptors with shortcuts, and `execute`, which runs in order with the pastes, cell entries and chart moves accepted before it, ends an IME composition and writes the text typed so far, and checks availability again, failing with `target-changed` when its selection moved meanwhile; print waits for the canvas to paint that text and fails with `render-failed` otherwise. The synchronous `api.save()` throws the new `XlsxSaveRefusedError` (`input-pending`, `input-failed`) instead of returning bytes without accepted input. A cell entry the engine refuses keeps the cell editor open with its error, and Save and Mod+S fail with `input-failed` until the entry is corrected or cancelled with Escape. Hosts compose built-in controls with their own actions through the `toolbar` and `showToolbar` props, `XlsxCommandProvider`, `useXlsxCommands`/`useXlsxCommandState`/`useXlsxCommand`, `EditorToolbar mode="commands"` (with `EditorToolbar.FormulaBar`), `ToolbarCommand`, `ToolbarCommandButton`, `ToolbarCommandSelect` and `ToolbarOverflow`, inside or outside the editor. The prop-based `EditorToolbar`, `Toolbar` and `useEditorToolbar` keep working in the default legacy mode. Toolbars, shortcuts and the proposals panel share one gate: read-only mode, cell selection, merge shape, collaboration and PNG support each disable commands with a coded reason. Shortcuts dispatch from the descriptors with platform-aware labels (now including Mod+B and Mod+I), only for the editor that owns the event, and leave text undo to the cell editor. Narrow toolbars move trailing groups into a keyboard-accessible More menu with radio choices, custom size and zoom prompts and color pickers, and `ToolbarDropdown` popups follow menu or dialog semantics by their content. New locale keys cover command labels and disabled reasons. The command and toolbar composition API is experimental and may change in minor releases.
- 667fae3: Host editor plugins in `XlsxEditor` through the new `plugins`, `pluginGrants` and `onPluginError` props. The plugin API is experimental and may change in minor releases. A plugin created with `defineXlsxPlugin` can contribute a docked panel (left, right or bottom of the grid, above the sheet tabs), an overlay on the grid, and commands registered as `plugin:<pluginId>/<id>`, whose results carry the plugin's own failure codes or a refused edit batch unchanged, with toolbar entries (`XlsxPluginToolbar` places them in replacement chrome) and shortcuts; `ToolbarCommandButton`, `ToolbarCommand`, `useXlsxCommand` and `useXlsxCommandState` accept those ids. Each activation receives `load`, `document-change` (the committed version after recalculation), `selection-change` (sheet, cells in their direction, chart), `mode-change`, `layout-change` and `grants-change` events, restricted clients (versioned `readCells`, `findText` and `validateEdits`, commands, granted edit batches, and `selectCells` and `scrollToCell` navigation that keeps focus unless asked), grid geometry that returns cell and range rectangles in overlay pixels for the painted frame, zoom, scroll and frozen panes, version-guarded state and `onCleanup` disposers. Plugins read, validate and navigate by default; built-in commands and edit batches need an explicit grant, rechecked with `readOnly` immediately before every write, and mutating built-in commands refuse plugins with `unsupported-policy`. Every contribution renders behind its own error boundary with a plugin-scoped command context and without the editor's formula bar binding, its keys and clicks never reach the grid, and a failing plugin is stopped and reported without affecting the editor or other plugins. Plugin shortcuts must use Mod or Alt, or a function key; built-in shortcuts win and a clashing plugin shortcut, including one the grid handles itself, is reported and not bound. The proposals panel now opens inside the grid area. New locale keys cover plugin panels, the plugin toolbar group and the new disabled reasons.

## 0.2.1

### Patch Changes

- 16d33a7: Fix locale declarations for TypeScript consumers with `skipLibCheck: false` and update React editors to depend on the corrected i18n packages.

## 0.2.0

### Patch Changes

- 93971b5: Remove outdated early-release warnings from package READMEs and link the JavaScript guide and changelogs.

## 0.1.0

## 0.0.8

## 0.0.7

### Patch Changes

- c6ad184: Add a Google Sheets-style toolbar to the XLSX editor backed by new engine
  APIs for range styling, number formats, selection-format aggregation, format
  painting, merge queries, and history state. Formatting is fully collaborative
  through a content-addressed style catalog (collaboration schema v3; v2 state
  does not migrate). Merging replaces intersecting ranges like Excel, parsing
  repairs overlapping merges in third-party files, and display-list font fields
  now serialize correctly so styled text renders with its real font, size, and
  weight.

## 0.0.6
