# @betteroffice/pptx-react

## 0.6.0

### Minor Changes

- e65df1c: Editors given `fonts` now load bundled faces for the other families a deck uses, including light members such as Calibri Light, so their text keeps its metrics. Supplied families keep the app's faces. Pass `defaultFonts={false}` to fetch no default fonts.
- 0979b79: Adds `title`, `description` and `decorative` to pptx read shapes, from shape alternative text at every depth. Sessions seeded before this change read them as absent until reseeded.

### Patch Changes

- 1d5efb7: The editors keep a left-to-right layout on right-to-left host pages, so the spreadsheet grid paints and menus stay on screen in Arabic or Hebrew apps.
- cc65f8e: READMEs state API compatibility as a versioning policy: releases are versioned with Changesets, the API can change between minor versions with every change in the changelog, and deprecations are announced before removal. The Python packages drop the development status classifier.
- fe7a975: Unconfigured editors now load bundled fonts per family before the first paint. App fonts still take precedence, and uncovered characters retain browser font fallback.
- 24179a4: Bold and italic runs in families whose bundled face has only a regular style, such as Chinese, Japanese and Korean fonts, are now drawn bold and slanted instead of upright regular. Lighter members such as Calibri Light keep their own weight.
- e67e831: The default PowerPoint editor downloads about two thirds less code up front: the deprecated legacy editor and the in-thread fallback editor now load only when used.
- 40f8868: Metafile pictures with labels, bitmaps or complex paint now draw through a lazy SVG replay in editors and native PNG exports, keeping the original media when saving. Adds `metafileSvg` and `metafileImageBytes` for hosts resolving display bytes.
- 01a0e1f: Plugin reads include `headerFooter` with visible slide footer, slide number and date text, plus stored settings in `layout.headerFooter`. These facts are included by default and optional for older engines.
- 9e8097c: The PowerPoint editor keeps its private crash-recovery copy of the opened file as a Blob instead of in page memory, lowering page memory by the file's size.
- eb1ac5d: Crash recovery in the PowerPoint editor now works on picture-heavy decks: it reopens the original file and replays the edits instead of loading an open-time state export that could exceed the 64 MiB update limit, and opening no longer spends time exporting that state.
- d4667b6: Lower PowerPoint editor memory on large decks: crash recovery keeps one copy of the opening state and the edits since then, instead of re-encoding and copying the whole presentation at every automatic checkpoint.
- fea115a: Slide thumbnails in the PowerPoint editor no longer keep their full-resolution pictures decoded after painting, lowering memory on picture-heavy decks.
- c485cb0: Typing no longer repaints the active slide's thumbnail on every keystroke; it updates once input pauses, and switching slides still shows each thumbnail's latest edits.
- ec768e6: Slide thumbnails in the PowerPoint editor decode pictures at thumbnail size instead of full resolution, lowering memory on picture-heavy decks.
- 61c7e5e: Typing in a PowerPoint text box no longer allocates and repaints a full slide canvas per keystroke: the editor reuses the slide raster and repaints only the lines that changed.
- c1d5485: Context menus, table tools, the shortcut dialog, rulers, error screens, presence chips and other labels that were always English now follow the editor locale, and numbers, units and times in them use its formatting.
- ad45f86: Editors decode packed fonts with their engine by default instead of downloading a separate decoder. Standalone font providers can opt into a decoder with `runtime.decoder`.
- 3d988e4: `XlsxEditor` and `PptxEditor` with `legacySyncApi` now call `onReady` once the first frame is painted. A hidden or zero-size editor calls it once it is shown.
- Updated dependencies [d17d538]
- Updated dependencies [5999962]
- Updated dependencies [f44c229]
- Updated dependencies [ad45f86]
- Updated dependencies [4fb3025]
- Updated dependencies [7b30d80]
- Updated dependencies [cc65f8e]
- Updated dependencies [842d624]
- Updated dependencies [1398cff]
- Updated dependencies [fe7a975]
- Updated dependencies [24179a4]
- Updated dependencies [e65df1c]
- Updated dependencies [84578e1]
- Updated dependencies [0530412]
- Updated dependencies [e664196]
- Updated dependencies [40f8868]
- Updated dependencies [4cfcf6e]
- Updated dependencies [01a0e1f]
- Updated dependencies [c45b6ff]
- Updated dependencies [eb1ac5d]
- Updated dependencies [d4667b6]
- Updated dependencies [8b61ea4]
- Updated dependencies [6329841]
- Updated dependencies [dfd1833]
- Updated dependencies [906e337]
- Updated dependencies [1398cff]
- Updated dependencies [0979b79]
- Updated dependencies [01a0e1f]
- Updated dependencies [ec768e6]
- Updated dependencies [76515fc]
  - @betteroffice/pptx@0.6.0
  - @betteroffice/pptx-i18n@0.6.0

## 0.5.1

### Patch Changes

- fa6fd11: Registers supplied fonts at their actual weight and style by default, allowing browser synthesis when only a regular face is available. Avoids duplicate registrations for the same family, weight, and style.
- 70e3441: Presentations open and show every slide when font downloads fail. Text falls back to available fonts and the editor logs one console warning.
- 375cbc3: Adds effective paragraph lists and slide titles, visibility and layout details to plugin content reads. Available by default with optional fields for compatibility with older engines.
- Updated dependencies [3d4b60c]
- Updated dependencies [baf21df]
- Updated dependencies [acbd0df]
- Updated dependencies [b1fb96b]
- Updated dependencies [27590d9]
- Updated dependencies [d9c0a82]
- Updated dependencies [70e3441]
- Updated dependencies [375cbc3]
- Updated dependencies [4eaba8b]
- Updated dependencies [59deeda]
- Updated dependencies [02818de]
  - @betteroffice/pptx@0.5.1
  - @betteroffice/pptx-i18n@0.5.1

## 0.5.0

### Minor Changes

- c688743: **Breaking:** Worker-only editing is now the default, and `onReady` receives `PptxAsyncEditorApi`. Until 0.6.0, the deprecated `legacySyncApi` keeps the synchronous API on the in-thread engine; `experimentalWorkerOpen={false}` selects the deprecated in-thread engine with the async API. See https://docs.betteroffice.dev/docs/migrating-to-0-5.
- c688743: **Breaking:** `onReady` receives `PptxAsyncEditorApi`: use `handleAsync()` instead of `handle` and `saveAsync()` instead of `save()`, and await `goToSlide`, `selectText`, `clearSelection`, `getPositionAtPoint`, `refresh` and `refreshProposals`, which now return promises.
- c688743: **Breaking:** `collaboration.onReplica` receives an `AsyncCollaborationReplica` from `@betteroffice/pptx`. Migrate to its promise-returning `encodeStateVector`, `encodeStateAsUpdate` and `applyUpdate`.
- 43f168f: **Breaking:** `DocxEditorProps`, `XlsxEditorProps` and `PptxEditorProps` are unions, so interfaces can no longer extend them and `Omit` or `Pick` keep only shared props. Extend `DocxAsyncEditorProps`, `XlsxAsyncEditorProps` or `PptxAsyncEditorProps` instead. See https://docs.betteroffice.dev/docs/migrating-to-0-5.
- c688743: **Breaking:** Collaboration refuses client ids of `Number.MAX_SAFE_INTEGER - 1` and above. Pin lower ids.
- 5cf8c6d: Adds `onDocumentChange`, `downloadOnSave`, `initialZoom`, `loadingIndicator` and `style` to `PptxEditor`, accepts an `ArrayBuffer` `file`, generates `collaboration.clientId` when omitted, adds `exportStructured()` and `exportMarkdown()` to the async editor API, and exports `PptxCommandAdmissionError`.
- c688743: The default editor supports plugins, notes, comments, proposal review, slideshows, exports, zoom, collaboration and `recoverySave`, and reports `onReady` failures through `onError`.
- 820feec: A presentation that fails to open, or an editor that fails to load, shows a styled, localized state with details; a failed load offers Reload page. Zoom spans 10% to 500%; Ctrl+Y redoes. Renamed test ids and attributes: `pptx-worker-editor`, `pptx-caret`, `pptx-caret-clip`, `pptx-preedit`, `data-pptx-resize`.

### Patch Changes

- 5cf8c6d: Fixes `onDocumentChange` and `onChange` notifications after API edits in the in-thread PPTX editor.
- 2bc343b: Worker editors now join collaboration rooms faster without visibly replaying each retained edit.
- e571025: An edit made just before the presentation worker fails now still reaches `onChange`, so hosts keep their unsaved-changes prompt after a crash.
- 098b2f7: Speaker notes typed just before or after switching slides now stay on the slide they were typed on, instead of being copied into the next slide's notes.
- df9c19c: Exporting a slide as PNG names the image after the presentation and slide number again (for example `Quarterly-slide-3.png`), and a failed export is reported through `onError`.
- 57887c4: `save()` on the default editor's `onReady` API now returns a promise with the presentation bytes, like XLSX and DOCX. `saveAsync()` keeps working as a deprecated alias.
- c688743: Fixes locked picture corner drags, typed toolbar fields, clicks on empty slide space and in table cells, Backspace and Delete on shapes, and typing into new text boxes.
- 2f9b322: The default editor centres the slide in the editing area again, with the white page and shadow of the previous editor.
- ba5ea2a: In the default editor, agent proposal cards name the slide and shape each change applies to, and slide thumbnails show their number beside the preview with a highlighted current slide, as before.
- 7bc019d: The worker-only editor shows the text cursor over slide text, the move cursor over shapes and the crosshair while an insert tool is armed, as the in-thread editor did.
- 9c6b82f: The default editor's agent proposals panel is docked beside the slide and styled like before, and the Add slide, Delete slide and speaker notes controls are styled again.
- 92f8e0a: The default editor and the read-only viewer show the Present button again. The slideshow starts on the current slide, and edits made while presenting show on the slide being presented.
- 136b29c: The default editor shows the agent proposals button again, with the number of pending proposals, and opens the proposals panel from it. The button is hidden in read-only viewers and when proposals are not available.
- 258cbb4: Runs collaboration inside the presentation worker by default for worker editors. Incoming update bursts merge together to keep the interface responsive and reduce repeated paints.
- fa13da2: In the worker editor, Tab and Shift+Tab inside a table cell select the next or previous cell's text, wrapping across rows. Tab in the last cell keeps the selection there; adding a row from it is not supported yet.
- Updated dependencies [43f168f]
- Updated dependencies [f18a42d]
- Updated dependencies [efa2982]
- Updated dependencies [c5f2b75]
- Updated dependencies [5cf8c6d]
- Updated dependencies [2bc343b]
- Updated dependencies [67c7645]
- Updated dependencies [c688743]
- Updated dependencies [c688743]
- Updated dependencies [c688743]
- Updated dependencies [c688743]
- Updated dependencies [820feec]
- Updated dependencies [c688743]
- Updated dependencies [fefa072]
- Updated dependencies [22f572a]
- Updated dependencies [1002e7e]
- Updated dependencies [5455c30]
- Updated dependencies [136b29c]
- Updated dependencies [258cbb4]
- Updated dependencies [fa13da2]
  - @betteroffice/pptx-i18n@0.5.0
  - @betteroffice/pptx@0.5.0

## 0.2.0

### Minor Changes

- ad64aa4: **Breaking:** collaboration updates now carry deck schema 2.2, and a 0.1.x update is migrated to 2.2 when it is opened. 0.1.x clients cannot open a document this version saved or migrated, so mixed-version collaboration across this upgrade is not supported: upgrade every client that shares a stored update.
- 7665432: Add version-checked, all-or-nothing PPTX edit batches. `PresentationHandle.version()`, `readContent()` and `findText()` return slides and story text with the session version they were read at; `validateEdits()` and `applyEdits()` resolve every step against that version and either commit text insertion, replacement, deletion and formatting, paragraph alignment, speaker notes, and shape rectangle, fill and outline changes as one transaction and one undo step, or return a typed refusal (`stale-version`, `missing-target`, `ambiguous-target`, `content-mismatch`, `overlapping-steps`, `unsupported`, `invalid-step`, `limit-exceeded`) with the deck untouched. `history: "none"` keeps a batch out of undo history, and `source` records provenance only. `PptxEditorApi` gains the same operations, flushing pending input first, refreshing the editor once after an applied batch, and refusing with `read-only` while the editor is read-only. The Rust `DeckSession` and `Presentation` and the Python `Presentation` expose the same API, and proposal previews and acceptance now share the batch staging. Ids and versions are session-scoped. The Rust `EditStep` and `EditFailureCode` enums are non-exhaustive.
- affeb36: Expose the PPTX editor's commands as `PptxEditorApi.commands`: serializable state with a stated reason for every disabled command, descriptors with platform-aware shortcuts, and `execute` that runs after pending input such as a decoding picture and checks availability again. Hosts compose built-in controls with their own actions through the `toolbar` and `showToolbar` props, `PptxCommandProvider`, `usePptxCommand`/`usePptxCommandState`, `EditorToolbar mode="commands"`, `ToolbarCommand`, `ToolbarCommandButton`, `ToolbarCommandSelect` and `ToolbarOverflow`, inside or outside the editor. The prop-based `EditorToolbar`, `Toolbar` and `useEditorToolbar` keep working and are deprecated. Canvas and panel proposal review, keyboard shortcuts and the default toolbar run through the same commands; a picture lands on the slide it was chosen for. Narrow toolbars move groups into a keyboard-accessible More menu, and `ToolbarDropdown` follows the menu or dialog pattern of its content. New locale keys cover command labels and disabled reasons. The command and toolbar composition API is experimental and may change in minor releases. `@betteroffice/pptx` adds `anchorCaret` and `resolveCaretAnchor`: plain-data caret positions that follow later edits, undo, redo and remote updates.
- 2eb2c55: Expose host save interception, awaited input flushing, and pointer position queries. Add manual undo capture, explicit boundaries, and undoable comment repositioning. `PptxEditorApi.save()` now throws while input or a pointer gesture is pending (await `flushPendingInput()` first) and after the presentation is replaced, instead of returning bytes that miss accepted input.
- 3d77b4f: Add host-controlled viewing mode, initial slide selection, and imperative slide and text navigation APIs.
- 911a294: Insert a picture onto a slide from the editor. The image mints its own media part, content-type default and relationship on save; `PptxEditor` gains a small "Insert image" icon button next to the text-box tool, and `PresentationHandle` gains `addPicture`. `addPicture` takes PNG, JPEG, GIF, BMP, TIFF, WebP and SVG images up to 8 MiB, checked before the picture reaches the deck, keeping oversized bytes out of collaboration updates.
- d11dff3: Host editor plugins in `PptxEditor` through the new `plugins`, `pluginGrants` and `onPluginError` props. The plugin API is experimental and may change in minor releases. A plugin created with `definePptxPlugin` can contribute a docked panel (left, right or bottom of the slide), an overlay on the slide canvas, and commands registered as `plugin:<pluginId>/<id>`, whose results carry the plugin's own failure codes or a refused edit batch unchanged, with toolbar entries (`PptxPluginToolbar` places them in replacement chrome) and shortcuts; `ToolbarCommandButton`, `ToolbarCommand`, `usePptxCommand` and `usePptxCommandState` accept those ids. Each activation receives `load`, `document-change`, `selection-change`, `mode-change`, `layout-change` and `grants-change` events, restricted clients (versioned `readContent`, `findText` and `validateEdits`, commands, granted edit batches, and `goToSlide`, `selectShape` and `selectText` navigation that keeps focus unless asked), slide geometry that converts slide EMU or pixels into overlay pixels and reports rendered shape bounds, version-guarded state and `onCleanup` disposers. Plugins read, validate and navigate by default; built-in commands and edit batches need an explicit grant, rechecked with `readOnly` immediately before every write, and mutating built-in commands refuse plugins with `unsupported-policy`. Every contribution renders behind its own error boundary with a plugin-scoped command context, and a failing plugin is stopped and reported without affecting the editor or other plugins. A superseded slide paint no longer draws over the slide painted after it. New locale keys cover plugin panels, the plugin toolbar group and the new disabled reasons.
- 911a294: Reorder a shape's paint order on its slide: bring to front, send to back, and step it forward or backward. `PresentationHandle` gains `bringShapeToFront`, `sendShapeToBack`, `bringShapeForward` and `sendShapeBackward`, and `PptxEditor`'s shape-formatting toolbar gains an "Arrange" menu for them.
- 8a5e6b8: Render PPTX pictures stored as SVG in both backends. The browser packages decode them with the browser's own SVG support, and the Rust and Python `render_png` paths rasterize them natively: shapes, paths, fills, strokes, gradients, clip paths, `<use>` and class-based stylesheets draw, under a sandbox that budgets the cost of expanding and painting a document before any of it runs. A tiled SVG repeats at its intrinsic size.

### Patch Changes

- 5114756: Drag and resize placeholders that take their geometry from the layout or master, from the frame they are drawn in, and draw them with the rotation and flips they inherit. The snapshot carries that geometry as `inherited`; the first `moveShape`, `resizeShape` or `setShapeRect` makes the whole transform the placeholder's own, so it keeps its size and orientation live and after a save. The Rust crates add `Placeholder::matches`, the placeholder matching that rendering and editing share, and the facade re-exports `InheritedGeometry`.
- 7faab61: A text box or shape drawn while a picture is still decoding now lands on the slide it was drawn on, even when you move to another slide before the picture finishes, and that slide's thumbnail shows it.
- Updated dependencies [6c30f4a]
- Updated dependencies [9639a6b]
- Updated dependencies [80341ac]
- Updated dependencies [6963a67]
- Updated dependencies [3fb2bf7]
- Updated dependencies [c02a145]
- Updated dependencies [030505a]
- Updated dependencies [ad64aa4]
- Updated dependencies [7665432]
- Updated dependencies [affeb36]
- Updated dependencies [8e8f97a]
- Updated dependencies [030505a]
- Updated dependencies [3e0c311]
- Updated dependencies [2eb2c55]
- Updated dependencies [5114756]
- Updated dependencies [911a294]
- Updated dependencies [030505a]
- Updated dependencies [6963a67]
- Updated dependencies [d11dff3]
- Updated dependencies [58f9bfb]
- Updated dependencies [60c79dd]
- Updated dependencies [7faab61]
- Updated dependencies [ad64aa4]
- Updated dependencies [7f158c7]
- Updated dependencies [3830d79]
- Updated dependencies [911a294]
- Updated dependencies [6963a67]
- Updated dependencies [27bf1fc]
- Updated dependencies [e0d12f3]
- Updated dependencies [18e1f32]
- Updated dependencies [73d405c]
- Updated dependencies [d76b4db]
- Updated dependencies [8a5e6b8]
- Updated dependencies [c5f1467]
- Updated dependencies [6c30f4a]
- Updated dependencies [ab3d722]
- Updated dependencies [030505a]
- Updated dependencies [af6292e]
- Updated dependencies [1f84618]
  - @betteroffice/pptx@0.2.0
  - @betteroffice/pptx-i18n@0.2.0

## 0.1.1

### Patch Changes

- 16d33a7: Fix locale declarations for TypeScript consumers with `skipLibCheck: false` and update React editors to depend on the corrected i18n packages.
- Updated dependencies [16d33a7]
  - @betteroffice/pptx-i18n@0.1.1
  - @betteroffice/pptx@0.1.1

## 0.1.0

### Minor Changes

- d6ba9da: Add session-local PPTX agent proposals with atomic acceptance, stale-target checks, rendered previews, and one-step Undo. Expose the workflow in Rust, TypeScript, and Python. Show inline text diffs and previous/proposed shape bounds on the React slide canvas, with proposal selection, accept/reject controls, and a review panel for before/after previews.
- 21b48f3: Edit per-slide speaker notes that persist through saves and collaboration updates, and present slides fullscreen from the React editor with keyboard navigation.

### Patch Changes

- 93971b5: Remove outdated early-release warnings from package READMEs and link the JavaScript guide and changelogs.
- bc34dfc: Parse connector shapes and preserve legacy collaboration updates when editing and saving.
- 69167fe: Paint justified lines at their caret positions and keep editor gestures consistent.
- b1f5c91: Render embedded TIFF pictures in browser presentations by converting them to PNG inside the PPTX WASM boundary. Uncompressed, LZW, PackBits and deflate sources are supported, including grayscale, RGB, palette and CMYK images; other compressions are skipped.
- 7ce54d6: Render bitmap-only WMF images in the slide editor and expose presentationImageBlob for canvas image resolvers. Preserve original media bytes when saving.
- Updated dependencies [93971b5]
- Updated dependencies [bd69e9e]
- Updated dependencies [d926fb0]
- Updated dependencies [cae162d]
- Updated dependencies [d6ba9da]
- Updated dependencies [6ae0b92]
- Updated dependencies [010865c]
- Updated dependencies [2877aba]
- Updated dependencies [1f30ea0]
- Updated dependencies [abb1e2c]
- Updated dependencies [899aac5]
- Updated dependencies [c4985a8]
- Updated dependencies [bfc3231]
- Updated dependencies [89a2134]
- Updated dependencies [d2aaf9c]
- Updated dependencies [c9b72bf]
- Updated dependencies [d6e6e91]
- Updated dependencies [bc34dfc]
- Updated dependencies [0c9b52e]
- Updated dependencies [69167fe]
- Updated dependencies [413499c]
- Updated dependencies [2044df7]
- Updated dependencies [8b48e8d]
- Updated dependencies [54fdaa0]
- Updated dependencies [cca2618]
- Updated dependencies [915dbaa]
- Updated dependencies [2b639b9]
- Updated dependencies [2c90c17]
- Updated dependencies [a61781d]
- Updated dependencies [25c7ea3]
- Updated dependencies [d280c87]
- Updated dependencies [22ce4e9]
- Updated dependencies [bf84789]
- Updated dependencies [3d95068]
- Updated dependencies [088d177]
- Updated dependencies [875d556]
- Updated dependencies [70e7394]
- Updated dependencies [acab663]
- Updated dependencies [0824bff]
- Updated dependencies [069e4d6]
- Updated dependencies [1e86217]
- Updated dependencies [89f8f7b]
- Updated dependencies [f5d9fd9]
- Updated dependencies [e5c4521]
- Updated dependencies [387f239]
- Updated dependencies [5c015e9]
- Updated dependencies [9274a2b]
- Updated dependencies [25d4ee4]
- Updated dependencies [7fdc0ee]
- Updated dependencies [21b48f3]
- Updated dependencies [ef5cdee]
- Updated dependencies [60113a3]
- Updated dependencies [253d680]
- Updated dependencies [a139ae9]
- Updated dependencies [051830e]
- Updated dependencies [07d72ce]
- Updated dependencies [1af946f]
- Updated dependencies [a3b2acd]
- Updated dependencies [2710a41]
- Updated dependencies [b1f5c91]
- Updated dependencies [bbd80c5]
- Updated dependencies [7ce54d6]
- Updated dependencies [863b70e]
  - @betteroffice/pptx@0.1.0
  - @betteroffice/pptx-i18n@0.1.0

## 0.0.4

### Patch Changes

- b962e66: Every OOXML chart family now draws with its own renderer instead of falling through to bars: area, scatter, bubble, radar, stock and surface join bar, line and pie. Stacked and percent-stacked grouping, gap width and overlap, marker symbols, data labels composed from `c:dLbls`, chart text from `c:txPr`, log scales, reversed axes, tick marks, gridlines and secondary value axes are all honoured, and `lumMod`, `lumOff` and `satMod` colour modifiers resolve so themed charts no longer draw oversaturated. Fixes horizontal bar charts, which ignored the zero baseline and drew negative values as nothing.
- 1b6a249: Charts in a presentation render for real instead of drawing a grey placeholder. Chart parts are loaded through the slide, layout and master relationship cascade, their colours resolve against the deck theme, and the plot streams into slide primitives with an accessible label. Data labels, axis titles and per-point colours draw, and an `ofPie` group now plots as a pie rather than as columns.
- 34541ae: PPTX decks now save with edits included, across every surface. The engine diffs the live CRDT state against a freshly seeded copy of the source package and writes back only what changed: untouched slides keep their exact source part bytes, edited slides are patched at the XML level so unmodeled markup — transitions, timing, unknown attributes — survives, and inserted or deleted slides rewrite `presentation.xml`, its relationships, and `[Content_Types].xml`. `Presentation::save()` in `betteroffice-pptx` no longer discards edits, `PresentationHandle.save()` returns the bytes on the npm core, `PptxEditor` gains a save toolbar button, Ctrl/Cmd+S, `onSave` and `fileName` props, and `save` on its `onReady` api, and the Python binding's `save`/`save_path` serialize edited decks instead of raising — `UnsupportedWriteError` is gone.

  Inside an edited paragraph, untouched runs keep their exact source markup; an edit contained in a single source run is rebuilt onto that run's properties, so hyperlinks, strikethrough, and spacing survive it. An edit spanning several source runs rewrites the span from the modeled styling — hyperlink and field bindings inside that span do not survive, which is the known write-back limitation.

- 6c7e94a: A deck snapshot persisted by the previous release opens again. Charts made the stored package a version 2 document, and the version check demanded an exact match, so every version 1 snapshot — every presentation a collaborator had already edited and saved — came back as `unsupported deck schema version` and could not be reopened.

  `open_from_update` now migrates instead of refusing. A version 1 document hydrates, its stored package is read back and rewritten in the current shape, and the document is stamped version 2, so the next snapshot the session writes is a version 2 one and the upgrade happens once. Nothing else in the document changes: the slide order, slide, shape and story containers were already identical between the two versions, and the only difference was the chart list the stored package gained. That list is optional when reading, so a package written before charts existed loads with none rather than failing on a missing field. Two clients opening the same old snapshot write the same migration and converge.

  A version this build does not know — a document from a newer release, or one whose version is missing or nonsense — is still rejected, and still reported before the stored package is parsed so the version is the error the caller sees.

- Updated dependencies [b962e66]
- Updated dependencies [1b6a249]
- Updated dependencies [6947366]
- Updated dependencies [34541ae]
- Updated dependencies [6c7e94a]
  - @betteroffice/pptx@0.0.4
  - @betteroffice/pptx-i18n@0.0.4

## 0.0.3

### Patch Changes

- 5212690: Google Slides-style editor toolbar for the PPTX editor: new-slide split button
  with layout picker, undo/redo, zoom, select and text-box tools, and contextual
  text formatting that also applies to whole shapes on selection. Text formatting
  now spans paragraph boundaries as a single undoable operation, double/triple
  click select word/paragraph, and roundRect corners render circular per the
  OOXML adj value instead of stretching with the shape.
- c134b2f: Collaborative presence: remote collaborators' shape selections render as colored outlines with name flags, with toolbar avatar chips and filmstrip dots showing which slide each peer is viewing.
- b87185f: Shape insertion and styling: a Slides-style shape picker inserts preset
  geometries (rectangles, ellipse, polygons, stars, arrows, chevron) by click
  or drag, and selected shapes get contextual fill, border color, border width,
  and corner-radius controls backed by new undoable, collaboration-native
  addShape/setShapeFill/setShapeStroke/setShapeAdjust engine operations.
- Updated dependencies [5212690]
- Updated dependencies [c134b2f]
- Updated dependencies [b87185f]
  - @betteroffice/pptx@0.0.3
  - @betteroffice/pptx-i18n@0.0.3

## 0.0.2

### Patch Changes

- 64e5940: Add pointer-based shape movement and text range selection to the PPTX editor.
- 69d62f1: Refine the XLSX and PPTX editor toolbars with compact DOCX-style control rails,
  grouped icon actions, and responsive value fields.
  - @betteroffice/pptx@0.0.2
  - @betteroffice/pptx-i18n@0.0.2
