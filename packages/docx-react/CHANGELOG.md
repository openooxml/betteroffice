# @betteroffice/docx-react

## 0.4.3

### Patch Changes

- 26e7db6: Plugin overlays get their positions when a document opened in a background tab is shown, without a scroll.
- 8a86196: Double-clicking a header or footer on any page of an editable document opens that header or footer for editing.
- 654df3c: With a header or footer open for editing, clicking another header or footer switches to it directly, without scrolling the page.
- f305310: With `experimentalWorkerOpen`, page counts stay whole while proposals are accepted, rejected or undone.
- Updated dependencies [d8fe3c0]
  - @betteroffice/docx@0.4.3
  - @betteroffice/docx-i18n@0.4.3

## 0.4.2

### Patch Changes

- 953cf7d: Double- and triple-click keep their selection when the pointer moves slightly.
- 3f09227: Plugin page overlays stay mounted while pages build in the background.
- 08c4763: Paragraphs lay out closer to Word: words across formatting runs, list numbers wider than their indent, `w:start`/`w:end` indents and bordered hanging indents. Picture outlines now paint.
- 195380d: While a long document lays out, pages that are not final yet show as placeholders instead of partial content.
- 260ba19: Read-only mode no longer opens header, footer or note editing.
- 4f10593: Repeated saves after adding a reply comment produce the same markup.
- ba93cf9: With `experimentalWorkerOpen`, read-only editors serve proposals, overlays, navigation and `search()` from the worker. New ref methods `getParagraphIdentities` and `resolveParagraphAnchors`; synchronous members that need the document throw `DocxReplicaNotReadyError` until it loads.
- 380dfac: Saving keeps what you didn't edit as it was, including inherited formatting, link targets and rich comments.
- 3b421e7: With `experimentalWorkerOpen`, layout stays in the worker after a synchronous ref call during load (a one-time warning names it), and a worker that fails after load is replaced; a repeated failure reaches `onError`.
- 5ff39d2: With `experimentalWorkerOpen`, "Page X of N" shows the saved page count until layout finishes.
- 864355f: With `experimentalWorkerOpen`, only pages near the viewport are prepared, so memory stays bounded on long documents; `whenLayoutComplete` also waits for the visible pages.
- 02c5459: With `experimentalWorkerOpen`, typing while an editable document loads is applied once it is ready.
- a0ea217: With `experimentalWorkerOpen`, the `previewFirstPage` preview runs in the worker and shows only final pages; `onFirstPagePainted` waits for one.
- bbfdeb2: With `experimentalWorkerOpen`, host proposals appear and decisions apply faster.
- Updated dependencies [08c4763]
- Updated dependencies [9c21cb8]
- Updated dependencies [4e6a9e4]
- Updated dependencies [195380d]
- Updated dependencies [4f10593]
- Updated dependencies [ba93cf9]
- Updated dependencies [380dfac]
- Updated dependencies [dcbc90b]
- Updated dependencies [4b5f430]
- Updated dependencies [ac0a55f]
- Updated dependencies [4797db4]
- Updated dependencies [1037096]
- Updated dependencies [5ff39d2]
- Updated dependencies [ed29b59]
- Updated dependencies [a0ea217]
- Updated dependencies [bbfdeb2]
  - @betteroffice/docx@0.4.2
  - @betteroffice/docx-i18n@0.4.2

## 0.4.1

### Patch Changes

- dab1fb3: Proposed changes, host edit batches and collaborators' updates are laid out in the background worker, so long documents stay responsive while they apply.
- ece75fd: Jumping to a proposal or paragraph no longer freezes long documents while the target is resolved. New `YrsSession.paragraphIdCount` counts the paragraphs of a story that carry an id.
- 2e160a2: Screen readers can read the text of every page again, including pages outside the visible window. Adds `reduceMirrorToText` and `buildMirrorPageText`.
- 90033fe: Saving no longer adds zero wrap distances or a zero effect extent to a picture that had none, so inline pictures keep their position in LibreOffice after a save.
- fbeac94: The first accept or reject of a proposal is laid out in the background worker like later ones, so it no longer freezes long documents.
- 85a0488: With `allowHostProposals`, clicking a host proposal no longer opens the built-in sidebar, and host proposals are left out of it unless `showHostProposalsInSidebar` is set.
- ddf51ae: With `experimentalWorkerOpen`, a read-only editor now loads its main-thread copy of the document only when something needs it, so a large document no longer freezes the page after it appears.
- 397bb00: Background page building is faster and keeps the editor responsive on documents with thousands of pages.
- e436875: Plugin overlays no longer vanish while an edit or proposal change lays out: they stay on the layout the pages still show until the new one arrives.
- 59ee81b: A picture inserted with the editor's Insert Image picker is now saved with its image, instead of a broken reference.
- 7460f7f: Accepting, rejecting or undoing a proposal no longer measures the whole document again when the decision shows or hides the only text in a font.
- 78a6325: Accepting, rejecting or undoing a proposal lays the document out in the resident worker only, instead of on the main thread first.
- a4c31e1: Saving from the editor keeps every part of the document you did not edit exactly as it was, so content the editor does not model survives the save.
- 9b7cbd5: Swapping the document on a mounted editor no longer opens the sidebar for re-proposed host proposals. Documents with their own tracked changes still open the sidebar automatically.
- db1a4b1: The sidebar no longer misses a new tracked change or comment that lands right after a document swap or relayout.
- 334b31f: Copying a table with extreme merged-cell spans no longer freezes the editor.
- 8231427: With `experimentalWorkerOpen` and `previewFirstPage`, the worker opens the full document while the preview paints, so large documents finish loading sooner.
- ff04916: The sidebar no longer re-reads a long document's unchanged paragraphs after each proposal or edit. `createYrsSidebarProjection` takes an optional `YrsStorySegmentSource` to read story segments through.
- ae7ffb2: Internal groundwork for keeping host proposals in the background worker while a read-only document opened there has no main-thread copy; nothing changes for editors yet.
- Updated dependencies [66d0d66]
- Updated dependencies [ff24d4d]
- Updated dependencies [9d8c52e]
- Updated dependencies [2624099]
- Updated dependencies [37a8ff8]
- Updated dependencies [36ab3fe]
- Updated dependencies [55d0103]
- Updated dependencies [4f2fcca]
- Updated dependencies [d83c2fd]
- Updated dependencies [81f0e0a]
- Updated dependencies [74322c9]
- Updated dependencies [dd21fb2]
- Updated dependencies [c58aebc]
- Updated dependencies [ece75fd]
- Updated dependencies [c3ebdd0]
- Updated dependencies [fd3f7d4]
- Updated dependencies [2e160a2]
- Updated dependencies [90033fe]
- Updated dependencies [e8fb199]
- Updated dependencies [a4c31e1]
- Updated dependencies [980b07b]
- Updated dependencies [a135ac1]
- Updated dependencies [540af05]
- Updated dependencies [eeffc30]
- Updated dependencies [7ab1ed4]
- Updated dependencies [28eaa43]
- Updated dependencies [ff04916]
- Updated dependencies [2988780]
- Updated dependencies [c407a61]
- Updated dependencies [570f8bf]
- Updated dependencies [ae7ffb2]
- Updated dependencies [7eeada9]
  - @betteroffice/docx@0.4.1
  - @betteroffice/docx-i18n@0.4.1

## 0.4.0

### Minor Changes

- 9ef6699: Add host proposals: `YrsSession.proposeChanges()` records a tracked-change batch outside undo history, and `setProposalStates()` previews accept/reject/restore via the new `revisionPreview` entry without changing the document. `getProposals()` and `YrsSession.onProposalChange()` read them; `DocxEditorRef`'s matching methods reach read-only viewers via the new `allowHostProposals` prop, off by default.
- b788552: Adds `search`, `searchNext`, `searchPrevious`, `searchGoTo`, `clearSearch`, `getSearchState` and `onSearchChange` to the editor ref, so a host's own find box can highlight and step through matches, also in a read-only editor. Find highlights now draw only for the pages in view.
- 73324c2: Add `getAnchorGeometry()`, page-aware plugin geometry for proposals, revisions, paragraphs, searches and text ranges. Proposal preview changes are now published and invalidate geometry, keeping overlays positioned as the document or preview changes.
- 0e3743c: Add the experimental `previewFirstPage` option to `DocxEditor`, off by default: it paints a read-only preview of a document's first pages while the whole document opens.
- 76026a5: Adds `withdrawProposals`, which settles finished host proposals as their decisions show, and `search: ''` proposals that fill an empty paragraph. Typing into an empty paragraph in the editor now takes the paragraph mark's formatting, as Word does.
- 6f0896c: Adds the opt-in `onFirstPagePainted` callback to `DocxEditor`, called once per document when its first pages are painted on screen.
- ae13248: Adds opt-in `experimentalWorkerOpen` for faster DOCX opening. Off by default.

### Patch Changes

- 66f7145: Pages stay painted, and scroll-to calls land on target, when a host scales the editor with CSS `zoom` on an ancestor. New `effectiveZoom(element)` export from `@betteroffice/docx/layout/render`.
- 6c1699f: `getAnchorGeometry` reads the document once per version and reuses those reads for later targets, so anchoring many proposals no longer rereads the whole document for each one.
- ef0bd77: Stop eagerly re-parsing the document after opening. The model for `onChange`/`onContentChange` now builds once the first pages appear, only if such a listener exists; save, export and `getDocument()` still build it on first use.
- 33e0e1d: `getCurrentPage()` now returns the page on screen at the moment it is called, and a read-only editor no longer scrolls back to its caret while pages build, so `scrollToPage` lands on its page.
- 62789d0: A document no longer fails to open with "resident pagination input is not built" when a layout reaches the display just as the editor switches to a new session, such as the handoff from a first-page preview.
- f47d11d: Editor teardown is more precise: a StrictMode remount keeps its worker, pages whose canvases no worker took repaint, a document that fails to open frees its session at once, and decoded images are released with their document.
- e4fd867: Highlights, remote cursors, table and image handles, link popups and comment cards now line up with the pages under a host's ancestor CSS `zoom`, and comment cards also at editor zoom other than 100%. `detectDisplayListTableInsertHover` takes an optional `buttonZoom`.
- a5b9424: Adds `preloadDocxEngine` and the `experimentalPrewarm` editor prop to prepare the editing engine ahead of opening a document. Both are opt-in.
- 3f1feb8: A superseded display-list query facade now answers from the live layout only within its own document line, never after the editor releases that document, and never across documents when no `line` is given. New `endDisplayListQueriesLine(line)` export from `@betteroffice/docx/layout/render`.
- a72d500: By default, editors sharing a page keep fonts apart: `onFontsLoaded` and `onError` ignore other editors' font loads, and same-named embedded fonts no longer override each other and are released on unmount. Direct callers opt in via `createFontLoadScope`, `registerDocumentFaces` and `loadEmbeddedFontFamilies`.
- add4cdb: Adds opt-in `set_windowed_incremental_builds` to limit incremental rebuilds to the display window and caret pages. The editor enables this for edits and proposal decisions, rebuilds other affected pages in the background, and waits for exact visible proposal geometry.
- 45e89f1: An editor that unmounts while its first layout is still running no longer throws "null pointer passed to rust" when that layout finishes.
- b626f7f: Adds `whenLayoutComplete()` to the editor ref, resolving with the page count once the whole document is laid out. Await it, not `onFirstPagePainted`, before reading page counts or content: until then `getTotalPages()` returns 0 and `getPageContent(n)` returns `null`, even for painted pages.
- 9c12794: The loading indicator stays centered over the editor until the first page paints, instead of jumping up when the toolbar and ruler mount; a custom `loadingIndicator` now shows too.
- 8a481c6: Embedded images stay compressed in the opened file until a page shows them. `createCanvasImageResolver` accepts a `media` source, and the new `mediaTokens` option (`DocxEditor`, `YrsOpeningOptions`), off by default, keeps images out of the shared document state.
- adb33ac: Selection and story queries on large documents no longer trigger whole-document rework; the first read of every story after an edit was quadratic in story count.
- 00f5a64: With opt-in windowed builds, full rebuilds build only visible pages and the caret's page. `set_display_retain_built_pages` keeps every built page, and the editor sets it while a page widget has focus.
- b7f564a: Page canvas lookups (pointer routing, caret, selection, highlight overlays, table handles, viewport anchoring) now read a registry instead of scanning the whole pages subtree.
- b6c407c: Large documents now keep each page's accessible text in the DOM only near the viewport, on by default. Tab, links, content controls and `RenderedDomContext.findElementsForRange()` still reach every page.
- 057dbf0: Plugin overlays, anchor geometry, sidebar cards and `RenderedDomContext` rectangles now line up with the pages when a host scales the editor with CSS `zoom` on an ancestor.
- 52040dd: A plugin's `scrollToParagraph` now waits for the layout to reach the target paragraph on large documents instead of failing after a second.
- 1e1c786: With both `previewFirstPage` and `experimentalWorkerOpen` on, the full document now opens in the worker after the preview paints, instead of on the main thread.
- 89c5d8e: Rebuild the editor's position projection incrementally: an edit now re-reads only changed paragraphs, via new `storiesChangedSince`, `storySegmentUnitDigests` and `storySegmentUnits` session queries.
- ea45cfd: `@betteroffice/docx-react` now re-exports the host-proposal types (`DocxProposalRequest`, `DocxProposalResult`, `DocxProposalStateRequest`, `DocxProposalSnapshot`, `DocxProposalFailure`, and related types), so hosts no longer need to import `@betteroffice/docx/yrs` directly.
- c5da06a: A read-only or viewing editor now selects text by dragging, double- or triple-clicking and Ctrl/Cmd+A, and copies it. Ctrl/Cmd+C copies the selection as plain text in every mode, with tabs, line breaks and tables kept.
- 9f65726: Loading another document into a mounted `DocxEditor` no longer sometimes replaces the editor with its error view and leaves the new document unlaid out.
- 6795634: Typing and deleting in table cells now goes through the resident engine worker in one request, like body paragraphs, instead of laying out the document first.
- f41899b: The resident worker now finishes a document's first layout in short steps, so requests that arrive meanwhile run between them, with identical results. `YrsSession.beginRegionLayout` and `resumeRegionLayout` expose the stepped region layout.
- 3f471de: Replacing or closing a document no longer throws "null pointer passed to rust" when the previous document's pages rebuild, or its fonts finish loading, after the switch.
- 0e8ca7e: Add `YrsRenderEnv.revisionPreview`, an opt-in map that lays out accepted or rejected tracked changes without changing the document. Plugin geometry follows a previewed decision once it is painted, and a paged export of a previewed layout is refused.
- c15d0ec: Reads a document's revisions and sidebar projection once per document version, so scrolling or building pages of a document with tracked changes no longer relists them.
- 0912663: Navigating to a distant page or position now jumps there directly, while nearby navigation still scrolls smoothly.
- cd2a6d7: The editor no longer loads, or reports errors for, East Asian or complex-script fonts a document names without having text in that script. `YrsDocxHost` gains `unusedScriptFonts`, which lists them.
- 4b39070: Suggested replacements now show the struck-out text before the new text, as Word does. `replaceRange` receipts carry the new text's `range`.
- cf221d8: Table resize handles are now built only for pages around the viewport, instead of every table on every page, via the new `deriveDisplayListTableFragmentsOnPages` query.
- dd9f00d: A new `tablePayload` session query reads just the current table's structure instead of the whole story, cutting per-keystroke cost in the toolbar and dialogs.
- 4538842: The viewport's scroll anchor now comes from the lines of on-screen pages instead of the whole document, via a new `visualLinesOnPage` query.
- 5c6e00c: Add `getMemoryStats()` and opt-in `onMemoryPressure` and `memoryBudget` for the editor's wasm memory. A worker that runs out of memory, or past the opt-in `memoryBudget.workerLimitBytes`, is replaced once, then reported as `ResidentWorkerOutOfMemoryError` without a main-thread fallback; other failures fall back as before.
- 4919560: The layout worker now starts while the document's fonts are still loading, so the first worker layout no longer waits for it to boot afterwards.
- 6809b31: Keep editing on the resident engine worker when a large document takes over five seconds to reply, instead of falling back to the main thread; only a worker silent for a minute is replaced. Queued keystrokes and worker-failure recovery now apply in order.
- Updated dependencies [f88ca8b]
- Updated dependencies [66f7145]
- Updated dependencies [726c5b4]
- Updated dependencies [81a2b31]
- Updated dependencies [c64ae45]
- Updated dependencies [5f5d4f0]
- Updated dependencies [b58c0b5]
- Updated dependencies [110414c]
- Updated dependencies [3d64f98]
- Updated dependencies [e4fd867]
- Updated dependencies [e49077d]
- Updated dependencies [a5b9424]
- Updated dependencies [3f1feb8]
- Updated dependencies [a72d500]
- Updated dependencies [384a90c]
- Updated dependencies [add4cdb]
- Updated dependencies [2763a8e]
- Updated dependencies [9ef6699]
- Updated dependencies [a6ac04e]
- Updated dependencies [ab5b7a8]
- Updated dependencies [86fcf80]
- Updated dependencies [0aa1c45]
- Updated dependencies [751515b]
- Updated dependencies [8a481c6]
- Updated dependencies [42ce38a]
- Updated dependencies [adb33ac]
- Updated dependencies [bf4471b]
- Updated dependencies [00f5a64]
- Updated dependencies [b204133]
- Updated dependencies [b7f564a]
- Updated dependencies [b6c407c]
- Updated dependencies [6ac2d8c]
- Updated dependencies [7ca1b0f]
- Updated dependencies [7205060]
- Updated dependencies [057dbf0]
- Updated dependencies [c4dbef0]
- Updated dependencies [89c5d8e]
- Updated dependencies [751515b]
- Updated dependencies [76026a5]
- Updated dependencies [7852788]
- Updated dependencies [77c07ec]
- Updated dependencies [6795634]
- Updated dependencies [f41899b]
- Updated dependencies [92cd82c]
- Updated dependencies [0e8ca7e]
- Updated dependencies [c15d0ec]
- Updated dependencies [4ad3ade]
- Updated dependencies [6a71241]
- Updated dependencies [d9472e7]
- Updated dependencies [85eac34]
- Updated dependencies [cd2a6d7]
- Updated dependencies [9b6cab0]
- Updated dependencies [0b8b954]
- Updated dependencies [4b39070]
- Updated dependencies [cf221d8]
- Updated dependencies [dd9f00d]
- Updated dependencies [dfa6b47]
- Updated dependencies [4538842]
- Updated dependencies [751515b]
- Updated dependencies [5c6e00c]
- Updated dependencies [2311375]
- Updated dependencies [91d414c]
- Updated dependencies [ae13248]
- Updated dependencies [92900b6]
- Updated dependencies [6809b31]
  - @betteroffice/docx@0.4.0
  - @betteroffice/docx-i18n@0.4.0

## 0.3.0

### Minor Changes

- 6c30f4a: TypeScript changes to check when upgrading from 0.2: parsed chart `ChartSeries.values`, `xValues` and `bubbleSizes` are `(number | null)[]`, with `null` for a point a sparse cache leaves empty; `YrsTableReceipt.changedStoryIds` is a required field, so receipt literals and mocks need it; `YrsResidentWorkerSnapshot.fonts` holds `YrsResidentFontRegistration` entries instead of raw bytes. In suggesting mode the editor now refuses edits it cannot record as tracked changes (inserting tables, page and section breaks, image layout and properties, page setup, watermarks, and table actions other than inserting and deleting rows), where 0.2 inserted tables as tracked rows and applied the rest untracked. Cmd/Ctrl+S now runs the editor's Save, including in read-only mode, and downloads the file unless `downloadOnSave` is `false`; hosts that handle the shortcut themselves take it over with `onSaveRequest` or by calling `preventDefault()` in a capture-phase listener.
- 9eb8d47: List DOCX content controls and fill text controls in version-checked edit batches. `YrsSession.listContentControls()` and `findContentControls(query)` return a session's controls in document order (body, headers, footers, footnotes, endnotes, comments; each control before the controls inside it) with the version they were read at: the structured export's control metadata (`controlId`, `ooxmlId`, `controlType`, `tag`, `alias`, `lock`, `showingPlaceholder`, `dataBound`) plus `placement`, `anchor`, `parentControlId`, the control's current text as `value` (or `unavailable` with a reason), `multiLine` and an `effectiveLock` that folds in the locks of the controls containing it. Tags, aliases and `ooxmlId`s match exactly, every match is returned, and `maxControls` and `maxBytes` refuse with `limit-exceeded` instead of returning part of the list; `complete: false` comes with diagnostics naming what could not be covered. `listDocxContentControls(bytes)` and `findDocxContentControls` read bytes without a session, `DocxEditorRef.listContentControls()` and `findContentControls()` flush pending input first, and the Rust `EditingDoc`, the `betteroffice-docx` `Document` (its current model) and the Python `Document` list the same controls.

  A `setContentControlText` batch step fills a plain- or rich-text control, addressed by `controlId` (the locator for the version it was read at), by a tag (the template author's name) or by `ooxmlId` (the authored `w:id`, which survives save and reopen), with plain text; a tag or `ooxmlId` must be carried by exactly one control of the document. The step replaces the inline control's content or the block control's paragraphs, clears the placeholder state in the typed flag, the captured `w:sdtPr` and the parsed properties, and keeps tags, aliases, ids, bindings and every other property as captured. LF is a line break, a plain-text control accepts it only with `w:multiLine` and keeps its lines in one paragraph, and a rich-text block control takes a paragraph per line; the text takes the first text run's formatting, or the control's own run properties while it shows its placeholder. Steps commit together as one undo step or not at all; refusals carry the batch code and a `reason` (`missing-tag`, `ambiguous-tag`, `missing-ooxml-id`, `ambiguous-ooxml-id`, `content-locked`, `bound-control`, `unsupported-children`, `nested-controls`, `multiline-not-allowed` and others). A header or footer part two relationships share is listed once and a fill writes both copies. Filling needs a session opened from DOCX bytes.

  Text controls now hold their text only as content. `setContentControlValue` with a string fills a text control through the same step, and every fill drops an authored value the control carried, outside undo history; embed and raw writes never give a text control an authored value, and retyping a valued control drops it. Collaboration updates still integrate whatever they carry: saving ignores a text control's value as before, and discovery reads the control's content and flags the value with `legacy-control-value`. An inline control is typed by its parsed properties before its flat attributes, as saving reads it. Checkbox, dropdown and date controls keep their typed setters. Parsing reads `w:text/@w:multiLine` instead of `w:val`, writes `w:multiLine` when it synthesizes plain-text properties, and keeps an inline control's `w:sdtEndPr`, which saving used to drop.

- 628cc9c: Add version-checked, all-or-nothing DOCX edit batches. `YrsSession.version()`, `readParagraphs()` and `findText()` return text with the session version it was read at, one U+FFFC per inline atom; `validateEdits()` and `applyEdits()` resolve every step against that version and either commit text insertion, replacement and deletion plus paragraph insertion, deletion and style changes as one transaction and one undo step, or return a typed refusal (`stale-version`, `missing-target`, `ambiguous-target`, `content-mismatch`, `overlapping-steps`, `locked-target`, `tracked-revision-conflict`, `unsupported`, `invalid-step`, `limit-exceeded`) with the document untouched. `history: "none"` keeps a batch out of undo history, and text steps can be recorded as tracked changes. `DocxEditorRef` gains the same four operations, flushing pending input first and refusing with `read-only` while the editor is read-only. Undo and redo now refresh every story they change before the next save. `proposeChange`, `addComment` and `applyFormatting` now resolve their targets in Rust, so text after a hard break, image, content control or note reference is addressed correctly. The Rust `EditingDoc` exposes the same API; its operation, target, failure-code and atom enums are `#[non_exhaustive]`.
- c28b2d4: Expose the DOCX editor's commands as `DocxEditorRef.commands`: serializable state with a stated reason for every disabled command, descriptors with shortcuts, and `execute` that runs after accepted input and checks availability again. Hosts compose built-in controls with their own actions through the `toolbar` prop, `DocxCommandProvider`, `useDocxCommand`/`useDocxCommandState`, `EditorToolbar` (now including `EditorToolbar.Review`), `ToolbarCommand`, `ToolbarCommandButton`, `ToolbarCommandSelect`, `ToolbarButton`, `ToolbarGroup`, `ToolbarSeparator` and `ToolbarOverflow`, inside or outside the editor. Dialogs and pickers a command opens apply to the document and selection they opened with (failing with `document-replaced` or `target-changed` otherwise), replacements and print wait for accepted input, and engine refusals fail the command. Narrow toolbars move groups into a keyboard-accessible More menu that keeps every choice of the built-in controls, and the keyboard help lists the command bindings. The selection context reports superscript, subscript and highlight, and `toggleMark` accepts superscript and subscript. New locale keys cover command labels and disabled reasons. The command and toolbar composition API is experimental and may change in minor releases.
- 7b6173f: Attach page fragments to DOCX structured exports. `exportDocxStructuredWithPages(bytes, options, { fonts, fontChains?, defaultChain, measurementDefaults?, renderEnvironment?, compatibility? })` takes plain JSON options with base64 font data, lays the bytes out in a private session with exactly those fonts in a measurement font store of its own, and returns the structured content with a deterministic snapshot page map; `YrsSession.exportStructuredWithPages(options)` reads the region layout a session retains and `DocxEditorRef.exportStructuredWithPages(options)` flushes input and exports the editor's layout of that version, computed from its current fonts, measurement defaults, render environment and pagination options. Page references describe that authoritative layout at the returned version. The map lists every physical page (blank parity pages included) with its section and displayed PAGE label, an occurrence for the body, each header or footer part and each note on every page that shows it, and fragments naming the exported block or inline, its paragraph-local text range or whole atom (marked `partial` when only part is on the page), table row windows with continuation and repeated-header flags, and optional geometry in unzoomed CSS pixels from the page's top-left corner. Ordinary exports never lay anything out. Paged exports refuse `stale-document`, `stale-layout` (section, settings or note metadata, fonts, options or `expectLayoutVersion`), `layout-unavailable` (including a font requirement with no registered font), `layout-not-converged`, `unsupported` and, while any laid-out story holds pending revisions, `unsupported-revision-layout` for the accepted and original views; `maxFragments` and `maxLayoutBytes` bound the map and mark it `truncated`. `renderDocxMarkdownWithPages` adds optional `<!-- docx-pages: -->` markers after each block marker and refuses a map from other content. The map carries its own `schemaVersion: 1`; additive fields keep it and a change existing readers would misread bumps it. In Rust, `EngineSession::export_structured_with_pages`, `export_structured_with_pages_for` and `export_snapshot_with_private_fonts` expose the same map.

  Page numbers of a section without `w:pgNumType w:start` now continue from the previous section instead of restarting at one, and PAGE fields show the continued number.

- 4884924: Add `onSaveRequest` for host-controlled File > Save and Cmd/Ctrl+S, plus awaited `flushPendingInput()` on the editor refs. Built-in saves flush pending input before serialization, and concurrent UI save requests share one workflow.
- 4cf2e55: Host editor plugins in `DocxEditor` through the new `plugins`, `pluginGrants` and `onPluginError` props. The plugin API is experimental and may change in minor releases. A plugin created with `defineDocxPlugin` can contribute a docked panel (left, right or bottom), an overlay positioned with `geometry.toOverlayRect`, which returns null once its layout is no longer rendered, sidebar cards anchored to versioned paragraphs, and commands registered as `plugin:<pluginId>/<id>`, whose results carry the plugin's own failure codes or a refused edit batch unchanged, with toolbar entries (`DocxPluginToolbar` places them in replacement chrome) and shortcuts; `ToolbarCommandButton`, `ToolbarCommand`, `useDocxCommand` and `useDocxCommandState` accept those ids. Each activation receives `load`, `document-change`, `selection-change`, `mode-change`, `layout-change` and `grants-change` events, a restricted read, command, edit and navigation client, version-guarded state and `onCleanup` disposers that run on removal, document replacement, revision change, unmount or failure. Plugins read, validate and navigate by default; built-in commands and edit batches need an explicit grant, rechecked with the editor mode and document policy immediately before every write, and mutating built-in commands refuse plugins with `unsupported-policy`. Every contribution renders behind its own error boundary with a plugin-scoped command context, and a failing plugin is stopped and reported without affecting the editor or other plugins. New locale keys cover plugin panels, the plugin toolbar group and the new disabled reasons. The snapshot-based `EditorPluginCore`, `PluginPanelProps`, `PanelConfig` and `SidebarItemContext` types in `@betteroffice/docx/plugin-api` are deprecated in favour of this API; `pluginOverlays`, `pluginSidebarItems` and `pluginRenderedDomContext` remain as deprecated unmanaged inputs.
- b153acd: Expose `RenderedDomContext.getPositionAtPoint()` for querying a client point without changing selection or focus, including page and story-region identity. The member is optional, so custom `RenderedDomContext` implementations written for earlier versions keep compiling; plugin and editor point queries answer `null` through a context that omits it.

  Add `DocxEditorRef.getPositionAtPoint()` and `DocxPluginGeometry.getPositionAtPoint()`, which return the hit with the document `version` its layout shows and a collapsed accepted-view `target` for edit batch steps, and `null` while input is pending or the painted layout is behind the document.

  Accept the hit result in `PagedEditorRef.displayPositionToYrsLoc()` to map body, header, footer, and note positions into their editing locations.

### Patch Changes

- 54e45fb: Save a comment moved with `setCommentRanges()` or a raw `setComment`, also into another story or by two collaborators at once, with one range and one reference at its new place, through undo, redo and collaborating replicas, instead of keeping the old reference beside a new one. A save writes one reference in each paragraph it rewrites that ends a comment's range, whatever order collaborators' changes arrived in, and drops that comment's references from the other paragraphs it rewrites while one remains; content kept as source bytes stays as it was.
- 3830d79: `saveYrsDocx` and the editor's Save write comment ranges that start or end inside a hyperlink, including one that holds a field or an equation, or inside a tracked change, or that follow a field, content control, equation, table or page break in their paragraph, as a paired start and end with a reference over the same text, and reopen them over that text. The editor's Save now also writes the ranges of comments added in the editor, in table cells and content controls too.
- ba7ca00: Keep DOCX paragraph identities across saves. Session keys and Word paragraph IDs (`w14:paraId`) are now separate: a session key is never saved, source IDs are kept as authored, and paragraphs authored in a session, edit batches included, get fresh, valid IDs that avoid every ID the package already uses. `saveYrsDocx()` returns the saved bytes plus a part-qualified persisted anchor for each saved paragraph, which `YrsSession.resolveParagraphAnchor()` resolves after reopening; session anchors resolve on every replica of one collaborative session, and each seeding open (`openDocx()`, `seedFromDocx()`, `documentToYrs()`) starts a new session so a stale anchor never resolves after reopening; a fixed `generation` option keeps shared seeds deterministic. Paragraphs whose ID the live session reassigned while a save ran are reported as `conflicts`. The React editor's Save writes and records the same IDs. `persistParagraphIds()` gives source paragraphs without an ID one on request across every story part, comments and note separators included, and refuses rather than guess at an ambiguous comment reference; a save whose stories are otherwise unchanged patches the IDs into the source bytes. `paragraphIdentities()` lists every paragraph's session, persisted and source anchors. Duplicate identities from concurrent edits or copies are repaired identically on every replica, source and saved IDs keeping theirs, and duplicate or malformed source IDs are no longer rewritten on parse; they save as authored until persisted. The Rust anchor-resolution, diagnostic and refusal enums are `#[non_exhaustive]`.
- ba2f5a7: Save no longer fails with "The document changed while saving" when the editor lays out again during the save; it still aborts when another document is loaded meanwhile. A failed input operation now fails only the flush, save or command that was waiting for it, so Save, print and toolbar commands keep working for the rest of the session.
- 25eca5d: Export DOCX as read-only structured content and Markdown. `exportDocxStructured(bytes, options)` returns schema version 1 JSON: ordered stories of paragraphs, headings with their outline-level source, list items with the rendered marker, tables on the source grid with spans, skipped grid columns and vertical-merge continuations, content controls, section breaks, and inlines for text, tabs, line, page and column breaks at their source positions, note and comment references, fields with their cached result read from the field's own runs, images (alt text and the owning part's relationship) and inline controls. Every block and inline carries an anchor (paragraph id, a batch-offset range in the view it names, table index, control id, a `sourcePart` location with the part's SHA-256 for retained XML, or `unlocated` with the reason content has no location of its own), and ids are deterministic. `revisionView` (`accepted`, `original` or `markup`) is required, and the markup view attributes insertions, deletions and moves; unreconstructed history becomes an anchored placeholder. Only the body is exported unless `stories` selects headers, footers, footnotes, endnotes or comments, and everything omitted or not represented is listed in `diagnostics`. `maxBlocks` and `maxBytes` stop the export at a whole block and set `truncated`. `exportDocxMarkdown` and `renderDocxMarkdown(content)` render Markdown, with entity-escaped HTML tables for complex tables and a `<!-- docx-export:N -->` marker per block mapped to its anchor. `YrsSession.exportStructured()` and `exportMarkdown()` read a live session and return the version the anchors belong to, changing nothing, and `YrsSession.headings(story)` lists a story's headings classified the same way. The Rust `EditingDoc`, the `betteroffice-docx` `Document` and the Python `Document` expose the same export; refusals are `invalid-options`, `limit-exceeded` or `unsupported` with `target: null`. JavaScript, React and the native editing session can attach page fragments from a configured layout.

  List numbering in sessions opened from DOCX now counts per list the way Word does: levels begin at their `w:start`, a numbering instance with a start override begins a list of its own, and deeper levels restart as `w:lvlRestart` says. Tracked insertions and deletions of images, shapes and charts stay on them in the editing session. The DOCX editor's outline reads headings from the engine, so document-default outline levels count and explicit body levels do not. `collectHeadings` is deprecated in favour of `YrsSession.headings()`.

- Updated dependencies [6c30f4a]
- Updated dependencies [9639a6b]
- Updated dependencies [8eddb2e]
- Updated dependencies [ca68499]
- Updated dependencies [6c30f4a]
- Updated dependencies [6c30f4a]
- Updated dependencies [6db9411]
- Updated dependencies [69b7cca]
- Updated dependencies [54e45fb]
- Updated dependencies [3830d79]
- Updated dependencies [432edfa]
- Updated dependencies [e62af65]
- Updated dependencies [9eb8d47]
- Updated dependencies [c087612]
- Updated dependencies [59d71c4]
- Updated dependencies [628cc9c]
- Updated dependencies [c28b2d4]
- Updated dependencies [7b6173f]
- Updated dependencies [0e9ca88]
- Updated dependencies [9b2fe7b]
- Updated dependencies [ebacc4e]
- Updated dependencies [9b2fe7b]
- Updated dependencies [72aed1b]
- Updated dependencies [b1d9284]
- Updated dependencies [6c30f4a]
- Updated dependencies [80432f3]
- Updated dependencies [e27dca8]
- Updated dependencies [ba7ca00]
- Updated dependencies [4cf2e55]
- Updated dependencies [b153acd]
- Updated dependencies [3c45811]
- Updated dependencies [023fad0]
- Updated dependencies [25eca5d]
- Updated dependencies [e4cb228]
- Updated dependencies [8273450]
- Updated dependencies [d655296]
- Updated dependencies [b46ad04]
- Updated dependencies [6c30f4a]
- Updated dependencies [4eb1d90]
- Updated dependencies [9b2fe7b]
- Updated dependencies [e151d79]
- Updated dependencies [2168e58]
- Updated dependencies [e5ad702]
- Updated dependencies [2958935]
- Updated dependencies [41f508b]
- Updated dependencies [1247270]
  - @betteroffice/docx@0.3.0
  - @betteroffice/docx-i18n@0.3.0

## 0.2.1

### Patch Changes

- 16d33a7: Fix locale declarations for TypeScript consumers with `skipLibCheck: false` and update React editors to depend on the corrected i18n packages.
- Updated dependencies [16d33a7]
  - @betteroffice/docx-i18n@0.2.1
  - @betteroffice/docx@0.2.1

## 0.2.0

### Minor Changes

- 295f42f: Keep one local undo history across document stories, group rapid keystrokes in WebAssembly, and preserve native undo in other inputs. Replace story-scoped history helpers with session-wide tracking and changed-story reporting.

  Migrate each removed API as follows: `historyStory()` returns the changed stories via `historyStories()` (sorted, empty before the first local edit instead of `null`); `undoDepth()` and `redoDepth()` are gone, query `canUndo()` and `canRedo()` instead; `markUndoGroup(startDepth)` is gone, rapid keystrokes now coalesce in WebAssembly with no host bookkeeping; `applyLocalUpdate(update, story)` drops its story argument and becomes `applyLocalUpdate(update)`; `beginUndoCapture(story, includeTableStories?)` drops its arguments and becomes `beginUndoCapture()`; `computeLayout()` no longer returns `blocks` and `measures`, read them lazily from `getLayoutKernelInputs(computation.layout)` as `measured` and `options`.

### Patch Changes

- 5069ad2: Keep accepted spreadsheet proposals undoable in collaborative sessions, preserve pending proposals through remote edits, and require a refreshed review when calculated previews change. Reject document suggestions that overlap partially tracked text. Existing public signatures and wire fields remain unchanged.
- 93971b5: Remove outdated early-release warnings from package READMEs and link the JavaScript guide and changelogs.
- 0b0a90a: Reuse successfully painted pages while scrolling and release temporary canvas buffers after replay.
- 43fad65: Reduce large-document interaction cost with per-line selection bands, lazy Unicode caret stops, compact retained-page shift replay, revision-bound lazy measured inputs, and stable page rendering identities.
- 1efec27: Improve DOCX fidelity with Word's 10 pt fallback for undeclared font sizes, short-paragraph widow control, and corrected table padding, minimum row heights, repeated headers, and rotated image sizing. Align automatic and wrapped tabs to the page grid, keep wrapped text metrics on their own lines, and hide list markers on page-break-only paragraphs. Preserve authored formatting and document state.
- 1d830df: Add a CDN-only font provider, settle Japanese font preflight without retry loops, and preserve floating header shapes without inflating body margins. Load and save alternate main-document filenames through their package relationships, and forward layout failures through the editor error callback.
- b1f5c91: Render embedded TIFF pictures in browser documents by converting them to PNG inside the DOCX parse WASM boundary. Uncompressed, LZW, PackBits and deflate sources are supported, including grayscale, RGB, palette and CMYK images; other compressions are skipped.
- 846b5d6: Fall back to the main-thread engine when the resident worker crashes, times out, or answers corruptly mid-input so typing survives worker failures. Replay the pending keystroke on the main-thread engine and keep genuine engine-level input rejections surfacing as errors.
- 0019657: Recover from resident worker crashes, WebAssembly traps, and unanswered requests so the editor can fall back to the main-thread engine. Reset retained worker frames and queries when switching engines so fresh main-thread frames render immediately.
- Updated dependencies [93971b5]
- Updated dependencies [b351bbe]
- Updated dependencies [2061849]
- Updated dependencies [4511b9a]
- Updated dependencies [4bf205b]
- Updated dependencies [0664bd3]
- Updated dependencies [5c04bc4]
- Updated dependencies [2c56acd]
- Updated dependencies [d4f4b85]
- Updated dependencies [a958376]
- Updated dependencies [cb2c6fe]
- Updated dependencies [356f67a]
- Updated dependencies [d4f4b85]
- Updated dependencies [2c658b6]
- Updated dependencies [297f43d]
- Updated dependencies [43fad65]
- Updated dependencies [6393137]
- Updated dependencies [4cc8dc0]
- Updated dependencies [1efec27]
- Updated dependencies [9281f7e]
- Updated dependencies [6779738]
- Updated dependencies [c1f9684]
- Updated dependencies [4d27f1e]
- Updated dependencies [b2ca63b]
- Updated dependencies [bd69e9e]
- Updated dependencies [20e913c]
- Updated dependencies [1d830df]
- Updated dependencies [6ce2439]
- Updated dependencies [e42ff8d]
- Updated dependencies [036f83e]
- Updated dependencies [a117530]
- Updated dependencies [9e4656f]
- Updated dependencies [1dc0e41]
- Updated dependencies [295f42f]
- Updated dependencies [56c3ca4]
- Updated dependencies [6f0e36d]
- Updated dependencies [284f0c4]
- Updated dependencies [d4f4b85]
- Updated dependencies [8adcd04]
- Updated dependencies [d4f4b85]
- Updated dependencies [b1f5c91]
- Updated dependencies [2ee434c]
- Updated dependencies [1a5ef23]
- Updated dependencies [1d0f41d]
- Updated dependencies [846b5d6]
- Updated dependencies [0019657]
- Updated dependencies [73cea54]
- Updated dependencies [d926fb0]
- Updated dependencies [c9b72bf]
- Updated dependencies [1e46a6f]
  - @betteroffice/docx@0.2.0
  - @betteroffice/docx-i18n@0.2.0

## 0.1.0

### Minor Changes

- 9540e23: Clicking a footnote or an endnote now opens it for editing. The caret lands under the pointer, typing goes into that note, and what is typed reaches the saved file. Escape leaves the note; clicking back in the body leaves it and places the body caret under the pointer.

  Undo follows the part being edited: the first edit in a newly opened part replaces the undo scope and discards the previous part's history.

  A note area is document text, not chrome, so a single click opens it — where a header or footer band, which repeats on every page and sits in the margin, still needs a double-click. The click that opens the note also places its caret: the position it resolved belongs to a story the editor is not on yet, so it waits for that story to become active and is discarded if the host never opens the note.

  The editor could previously be in a header/footer band, and now also in a note, so the two are one value rather than a flag each: `partEdit` names the single non-body part that is open, and opening a note closes whatever was open before. That removes the state where a band and a note could both claim the caret, and folds the band's first-page variant and originating page into the same value. `PagedEditor` takes `partEdit` in place of `hfEditMode` / `hfEditRId` and reports the open part's selection through `onYrsPartSelectionChange`.

  Everything the body offers behind an open band was already off, and stays off behind an open note for the same reason plus one of its own: the display list indexes images, tables and hyperlinks by band, so a note area has no scope to look one up in and answering `body` would hand back something from behind the note. Selecting a picture, dragging a table edge, the row/column insert affordance and opening a link are therefore inert inside a note; caret movement, word and paragraph selection, drag-selection and Home/End are not. Arrows navigate a note the way they navigate a band — paragraph by paragraph, since the display-list vertical-move query answers body positions only — and the caret never leaves the note it is in.

  Selection geometry comes from the note's own story: `noteCaretRects` is the note twin of `hfCaretRects`, resolving the leading edge of the caret's position and falling back to the trailing edge of the previous one at end of line. A note paints on exactly one page, so unlike a band there is never a second candidate page to choose between. One overlay now paints the caret and highlight for whichever part is open, band or note, and it asks the queries directly — `computeHfCaretRectsFromDisplayList` and `computeHfSelectionRectsFromDisplayList`, which had that overlay as their only caller, are gone from `@betteroffice/docx/layout`.

- 9540e23: Header/footer editing state becomes one value instead of a flag per kind. `partEdit` names the single non-body part the editor has open, folding the band's first-page variant and the page it was opened from into it — three `useState`s in `DocxEditor` become one.

  `PagedEditor` takes `partEdit` in place of `hfEditMode` / `hfEditRId`, and reports the open part's selection through `onYrsPartSelectionChange` in place of `onYrsHfSelectionChange`. `CanvasHfSelectionOverlay` becomes `CanvasPartSelectionOverlay` and queries the display list directly, so `computeHfCaretRectsFromDisplayList` and `computeHfSelectionRectsFromDisplayList` are gone from `@betteroffice/docx/layout`.

  Behaviour is unchanged, with one exception: Escape moves from the header/footer chrome component to the paged area, so it now also closes a band whose chrome never mounted.

- 6be0c18: Bundled metric-compatible fonts ship as `@betteroffice/fonts`, plus `@betteroffice/fonts-cjk` for Chinese, Japanese or Korean, and DOCX uses them only when you hand the module over: `configureDefaultFonts({ fonts })`, or `configureDefaultFonts({ load: () => import('@betteroffice/fonts') })` to keep it in its own chunk. Installing the packages alone does nothing — without that call the engine reaches for no font package, measurement falls back to the browser, and pagination will not match Word. Because `@betteroffice/docx` no longer names `@betteroffice/fonts` anywhere in its published bundle, an esbuild consumer without the optional peer builds again.

### Patch Changes

- e88e5e7: Typing and deletion in the DOCX editor now follow the visible caret after tables, page or column breaks, and block-level content controls.
- 8aa8b37: Edits made in an endnote now reach the saved file. Endnote stories were seeded into the collaborative document but never projected back on save, so every endnote edit was silently replaced by the imported text; footnotes and endnotes now project through one path, and typing in a note or header — including a table cell inside one — marks its root story for the changed-stories-only save.
- 5ff0142: Paragraphs with `lineRule="exact"` or `lineRule="atLeast"` now use Word's measured baseline model
- a2d08a8: The DOCX editor's mouse cursor now reflects what is under it. Hovering typeable text shows the caret cursor and everything else shows the arrow, where the canvas renderer previously painted one arrow over the whole document.

  A position alone could not drive this: the display hit test resolves a caret everywhere on a page that carries any text, margins included, so it cannot say whether a click there would type. `hit_test_regions` now also reports what the point landed on, as `target` on its result — `"text"`, `"image"` or `"none"`. It is the same answer a click acts on and follows the same order: a selectable picture first, since the pointer path picks one before it ever asks for a position, then a run's own box, then the page's typeable area. That area is the authored content box, so the gutter between columns counts like the columns it separates, minus the page's note areas, which this path cannot edit — a click in a footnote lands in the body above it, and the pointer no longer invites one. An area with no positionable text reads as no target for the same reason: a click there jumps the caret to the end of the document.

  A picture is a select target rather than text whatever its wrap mode, because both inline and anchored images carry a document position. One that carries none cannot be selected — a picture watermark — so text painted over it stays readable through it.

  `target` is additive and optional on `DisplayListRegionHit`, so a display-list query answered by an older wasm build still typechecks and reads as "not text". All three query paths — the stateless JSON exports, the session handle, and the resident editing engine — answer from the same resolver, so none can disagree.

  Header and footer bands read as text over their own runs and arrow elsewhere: a double-click there opens the band for editing at that word. While one is open only that band types; a read-only document types nowhere.

- ba687e8: Footnote and endnote text is now reachable by the display hit test. A point that lands in a note area resolves against that note instead of the body line above it, and the pointer shows the caret cursor over a note's glyphs where it used to show the arrow.

  A note is a document of its own, so the region vocabulary had to grow to name one. `hit_test_regions` answers `"footnote"` / `"endnote"` alongside `"body"`, `"header"` and `"footer"`, carrying `noteId` the way a band carries its `rId`: the position then addresses the `fn:{id}` / `en:{id}` story, never the body. A note area stacks several such stories, so the point resolves against the note nearest it rather than the area as a whole — a click can never borrow a position from another note. Nearest counts both axes, because a note area lays its notes out in columns it starts at the same vertical position: telling a pointer in one column from the next takes more than its `y`. An area that paints nothing still owns the click, like an empty header band does, and names no story to route it to.

  That made the display list's own positions load-bearing. A note's primitives kept the note story's range, but anything without one inherited the body range of the reference mark anchoring the note — and under a note region that body number reads as a note-story position, a different document entirely. The reference label leading every note is exactly such a primitive, so this was no edge case: hovering a footnote's number answered with a body position. Those primitives stay unpositioned, and the body anchor rides on the region's `notes` entry, where the accessibility mirror already reads it.

  The label being presentation rather than story content has a visible consequence: it reads as no target, so the cursor is an arrow over a note's number and a caret over its text. A band behaves the same way over its own non-text, and the rule the cursor follows is to under-claim rather than promise typing where a click will not.

  Selection geometry follows the same scoping. `range_rects_in_region` takes the region and the part that owns it as one argument, so a header/footer is named by an `rId` and a note by an id that is never optional — two notes are two unrelated documents whose positions must not mix. `noteRangeRects` is its query-facade twin, and a selection over a note that runs to a bordered paragraph or a table keeps every line it covers rather than stopping at the border.

  The test that pinned note areas as holes in the typeable area is replaced by tests of the new behavior, and the area subtraction it described is gone: a point inside a note is answered by the note before the body is ever asked.

  `region` may now be `"footnote"` or `"endnote"` and `noteId` is optional on `DisplayListRegionHit`, so a display-list query answered by an older wasm build still typechecks. Clicking a note still does nothing — routing a selection into a note story is the editing mode's job, not the layout API's.

- 3533411: The table border toolbar's buttons now do what their names say. `set_cell_borders` replaced a cell's whole `tcPr.borders` object with whatever it was handed, so pressing Top Border on a cell of a bordered table silently deleted that cell's other three rules, and Outside Borders gave every selected cell all four edges — a full grid rather than an outline. It now merges the sides it is given: an omitted side is left alone, `style: "none"` authors an explicit no-border, and a JSON `null` drops the authored side, the same patch convention `set_cell_text_format` already uses.

  Inside Borders wrote `insideH`/`insideV` straight onto each cell. Those keys describe a table's interior edges and have no meaning on a single cell, so nothing rendered them and the grid vanished on screen, while the writer still lifted a complete `w:tblBorders` into the file and the rules reappeared on reopen. `insideH`/`insideV` now resolve per cell to the physical edges interior to the selection, reusing the mapping seeding already applies when it pushes `w:tblBorders` down to cell positions — so an Inside Borders command paints the interior rules of whatever is selected and leaves the outline untouched, and a single-side button applies to the selection's own edge rather than to every cell in it.

  Saving no longer invents table borders no cell carries: the `w:tblBorders` lifted back out of a table is now read from the cells that own each boundary, with `insideH`/`insideV` taken from an interior edge, instead of filling every missing side from the first border found anywhere in the table.

- Updated dependencies [b962e66]
- Updated dependencies [53c583d]
- Updated dependencies [f6af707]
- Updated dependencies [8aa8b37]
- Updated dependencies [5ff0142]
- Updated dependencies [c2e9e69]
- Updated dependencies [a2d08a8]
- Updated dependencies [17f2ead]
- Updated dependencies [7799555]
- Updated dependencies [9540e23]
- Updated dependencies [ba687e8]
- Updated dependencies [43ab7ba]
- Updated dependencies [9540e23]
- Updated dependencies [335bb21]
- Updated dependencies [3533411]
- Updated dependencies [53c583d]
- Updated dependencies [d2e9577]
- Updated dependencies [cd305a5]
- Updated dependencies [6be0c18]
  - @betteroffice/docx@0.1.0
  - @betteroffice/docx-i18n@0.1.0

## 0.0.4

### Patch Changes

- 5c9a482: ArrowUp/ArrowDown move the caret by visual line with persistent goal-X (including across paragraphs, pages, columns, and into tables), and content below tables is clickable and editable again.
- 5c9a482: Collaborative presence: remote collaborator carets and selections render as colored overlays with name flags, anchored by yrs sticky indices so they rebase exactly under concurrent edits; carets follow remote typing instantly by inferring position from document updates.
- 5c9a482: Opening a document now seeds the collaborative session directly in the Rust engine instead of materializing the full TypeScript document model and projecting it; the TS model is built lazily only where the public API still exposes it, and the internal DrawingML host package is dissolved.
- 5c9a482: Remote collaborators' edits no longer move the local viewport: relayouts triggered by remote updates anchor to the topmost visible line via yrs sticky positions and compensate the scroll offset, while caret scrolling fires only for local actions. Anchoring holds across page boundaries too, so text overflowing onto a new page (or pulling back off one) no longer jumps the viewport for either the author or a viewer.
- Updated dependencies [5c9a482]
- Updated dependencies [5c9a482]
- Updated dependencies [5c9a482]
- Updated dependencies [5c9a482]
  - @betteroffice/docx@0.0.4
  - @betteroffice/docx-i18n@0.0.4

## 0.0.3

### Patch Changes

- Updated dependencies [b34bb01]
  - @betteroffice/docx@0.0.3
  - @betteroffice/docx-i18n@0.0.3

## 0.0.2

### Patch Changes

- eed05a6: Fix the published dependency ranges: 0.0.1 shipped the unresolved `workspace:*` protocol for `@betteroffice/docx` and `@betteroffice/docx-i18n`, which made `npm install @betteroffice/docx-react` fail. Ranges are now pinned to concrete versions at publish time.
  - @betteroffice/docx@0.0.2
  - @betteroffice/docx-i18n@0.0.2
