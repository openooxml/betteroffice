# @betteroffice/docx-i18n

## 0.5.1

## 0.5.0

### Minor Changes

- 43f168f: **Breaking:** A complete `LocaleStrings` object needs the new keys, including the PPTX `worker` group. `Translations` and `PartialLocaleStrings` are unaffected.

### Patch Changes

- 3690563: A document that cannot open now shows a short explanation, with the technical message under Show details. When the editor itself fails to load, a Reload page button fetches it again. A failed save reports a plain reason instead of internal text.
- e030e1d: Tracked-change cards are one compact row: author icon, name, time and icon-only accept and reject buttons. Adds `trackedChanges` with an opt-in `display: 'dots'` margin mode, `getAuthorAppearance` and `renderCard`.
- a205b2a: Typing into a document that is still opening no longer disappears silently: the editor says typing is paused until the document is ready, and leaves the document unchanged.
- 3690563: Worker editor labels, context menu items, object dialogs, note buttons, the draft recovery notice and the document reader now follow the editor locale. New keys have English values; other locales fall back to English for them.

## 0.4.3

## 0.4.2

## 0.4.1

## 0.4.0

## 0.3.0

### Minor Changes

- c28b2d4: Expose the DOCX editor's commands as `DocxEditorRef.commands`: serializable state with a stated reason for every disabled command, descriptors with shortcuts, and `execute` that runs after accepted input and checks availability again. Hosts compose built-in controls with their own actions through the `toolbar` prop, `DocxCommandProvider`, `useDocxCommand`/`useDocxCommandState`, `EditorToolbar` (now including `EditorToolbar.Review`), `ToolbarCommand`, `ToolbarCommandButton`, `ToolbarCommandSelect`, `ToolbarButton`, `ToolbarGroup`, `ToolbarSeparator` and `ToolbarOverflow`, inside or outside the editor. Dialogs and pickers a command opens apply to the document and selection they opened with (failing with `document-replaced` or `target-changed` otherwise), replacements and print wait for accepted input, and engine refusals fail the command. Narrow toolbars move groups into a keyboard-accessible More menu that keeps every choice of the built-in controls, and the keyboard help lists the command bindings. The selection context reports superscript, subscript and highlight, and `toggleMark` accepts superscript and subscript. New locale keys cover command labels and disabled reasons. The command and toolbar composition API is experimental and may change in minor releases.
- 4cf2e55: Host editor plugins in `DocxEditor` through the new `plugins`, `pluginGrants` and `onPluginError` props. The plugin API is experimental and may change in minor releases. A plugin created with `defineDocxPlugin` can contribute a docked panel (left, right or bottom), an overlay positioned with `geometry.toOverlayRect`, which returns null once its layout is no longer rendered, sidebar cards anchored to versioned paragraphs, and commands registered as `plugin:<pluginId>/<id>`, whose results carry the plugin's own failure codes or a refused edit batch unchanged, with toolbar entries (`DocxPluginToolbar` places them in replacement chrome) and shortcuts; `ToolbarCommandButton`, `ToolbarCommand`, `useDocxCommand` and `useDocxCommandState` accept those ids. Each activation receives `load`, `document-change`, `selection-change`, `mode-change`, `layout-change` and `grants-change` events, a restricted read, command, edit and navigation client, version-guarded state and `onCleanup` disposers that run on removal, document replacement, revision change, unmount or failure. Plugins read, validate and navigate by default; built-in commands and edit batches need an explicit grant, rechecked with the editor mode and document policy immediately before every write, and mutating built-in commands refuse plugins with `unsupported-policy`. Every contribution renders behind its own error boundary with a plugin-scoped command context, and a failing plugin is stopped and reported without affecting the editor or other plugins. New locale keys cover plugin panels, the plugin toolbar group and the new disabled reasons. The snapshot-based `EditorPluginCore`, `PluginPanelProps`, `PanelConfig` and `SidebarItemContext` types in `@betteroffice/docx/plugin-api` are deprecated in favour of this API; `pluginOverlays`, `pluginSidebarItems` and `pluginRenderedDomContext` remain as deprecated unmanaged inputs.

## 0.2.1

### Patch Changes

- 16d33a7: Fix locale declarations for TypeScript consumers with `skipLibCheck: false` and update React editors to depend on the corrected i18n packages.

## 0.2.0

### Patch Changes

- 93971b5: Remove outdated early-release warnings from package READMEs and link the JavaScript guide and changelogs.

## 0.1.0

## 0.0.4

## 0.0.3

## 0.0.2
