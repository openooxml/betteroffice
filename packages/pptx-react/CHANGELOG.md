# @betteroffice/pptx-react

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
