# @betteroffice/pptx

## 0.1.0

### Minor Changes

- d6e6e91: Read, add, reply to, resolve and remove PowerPoint comments in both the legacy and modern formats, saved by patching the existing comment XML in place (deck schema v8; older clients reject new updates until upgraded).
- bf84789: Set paragraph alignment from the presentation toolbar.
- 387f239: Add PPTX-to-PNG export for Rust, Python, and browsers, plus an Export PNG button
  in the React editor.

  Rust: enable the `raster` feature; exhaustive `Error` matches must handle the new
  `Raster` variant.

### Patch Changes

- cae162d: Drop redundant buffer copies around the wasm boundary and per collaboration update.
- 010865c: Measure block arrow heads from the shortest side while preserving shaft widths across aspect ratios.
- 2877aba: Render automatic list numbers with bullet formatting, preserve explicit restarts when editing and saving, and migrate collaboration snapshots to schema 11 after the schema-10 baseline migration.
- 1f30ea0: Apply picture duotone, biLevel, grayscale and colour-change effects in Canvas, PNG exports and the native viewer, preserving alpha and migrating older collaboration snapshots.
- 899aac5: Round an automatic value axis to whole `{1, 2, 5} x 10^k` steps and widen its unpinned ends to the next step, so a stacked bar no longer ends on the plot edge.
- c4985a8: Respect explicit chart data label settings that disable every field.
- bfc3231: Reserve space for top and bottom chart legends, wrap their entries to fit, and center PowerPoint chart titles.
- d2aaf9c: Paint a chart's own `c:chartSpace` fill instead of a white ground, and stroke each axis line with its own `a:ln` colour and width, or not at all under `a:noFill`.

  Migrate collaboration snapshots to schema 15 after the existing schema-14 gradient-outline migration, importing chart-space fills and axis lines from a reattached source.

- bc34dfc: Parse connector shapes and preserve legacy collaboration updates when editing and saving.
- 0c9b52e: Render numeric custom geometry paths, including elliptical arcs and separate path fills and strokes. Preserve custom geometry in schema 7 collaboration snapshots and migrate versions 1–6 after the existing hidden-shape migration.
- 69167fe: Paint justified lines at their caret positions and keep editor gestures consistent.
- 413499c: Preserve weight and slant when substituting missing presentation fonts, and choose
  the nearest fallback style consistently regardless of face registration order.
- 2044df7: Render gradient outlines across browser, raster, and native backends while preserving their paint through theme inheritance, width edits, and legacy snapshot migration.
- 8b48e8d: Render unordered gradient stops correctly in slide display lists and raster output while preserving equal-position stop order and source XML.
- 54fdaa0: Skip hidden slide shapes and hidden groups' descendants when painting and hit-testing.

  Deck schema 6 migrates existing version 1–5 documents by recovering hidden flags from stored package data after the schema-5 theme-formatting migration. Older clients reject the new schema.

  `ShapeSnapshot.hidden` is optional and omitted when false, preserving unchanged snapshot JSON. Only hidden shapes store a Yrs key; the schema stamp changes for all decks.

- cca2618: Transpose horizontal bar chart axes, preserve category direction, and reserve space for category labels, axis titles, and secondary value ticks.
- 2b639b9: Render triangle, open arrow, stealth, diamond, and oval line ends on the PPTX
  canvas, preserving their independent width and length settings.
- 2c90c17: Apply inherited list styles and render character bullets with their own formatting while preserving caret positions and migrating collaboration snapshots to schema 9.
- 3d95068: Apply inherited paragraph line spacing, preserve baselines for expanded point spacing, and migrate collaboration snapshots to schema 12 after baseline and numbering migrations.
- 70e7394: Crop pictures to their `srcRect` and clip and outline their preset masks in Canvas, PNG exports, and the native viewer. Preserve JSON for uncropped rectangular pictures.
- 0824bff: Scale chevron and homePlate adjustments from the shortest side, allow their points to span the full width, and normalize DOCX preset guide values consistently.
- 069e4d6: Render superscript and subscript runs at their shifted baselines and preserve their formatting through edits and saved decks.
- 1e86217: Render gradient-filled text using its lowest valid stop and preserve authored gradients across text edits until their color changes.
- 89f8f7b: Preserve each text run's colour, weight, italic, underline, and size while keeping identically styled adjacent runs grouped.
- 9274a2b: Render stretched picture fills through shape geometry and retain their source data across collaboration snapshots.

  Migrate collaboration snapshots to schema 13 after the existing schema-12 line-spacing migration, preserving source imports and edited text.

- 25d4ee4: Use shape font-reference colours above master text defaults while preserving run, paragraph, and placeholder colours.
- 7fdc0ee: Evaluate slide-number fields on masters and layouts, counting from the presentation's first slide number.

  Collaboration snapshots use deck schema v4. Older snapshots migrate through the
  existing v3 connector migration before the v4 slide-number migration; missing
  starting numbers default to one. Readers supporting only v3 reject v4 snapshots.

- 07d72ce: Keep shape text unmirrored and honor vertical text direction, insets, and caret positions without changing ordinary horizontal rotations.
- 1af946f: Render overflowing text at its intended size and anchor across backends, preserve explicit clipping, and keep transformed text clickable while migrating collaboration snapshots to schema 16.
- a3b2acd: Resolve DrawingML font references to the theme's major or minor script face, using its Latin face when the requested script slot is empty.
- 2710a41: Resolve PPTX theme fill and line references, including background fills and placeholder colour transforms. Preserve explicit shape and placeholder formatting. Preserve font reference colours from the existing text-style resolver.

  Migrate v1–v4 deck snapshots to schema 5 after the existing connector and slide-number migrations, preserving edits, numbering and source ordinals.

## 0.0.4

### Patch Changes

- b962e66: Every OOXML chart family now draws with its own renderer instead of falling through to bars: area, scatter, bubble, radar, stock and surface join bar, line and pie. Stacked and percent-stacked grouping, gap width and overlap, marker symbols, data labels composed from `c:dLbls`, chart text from `c:txPr`, log scales, reversed axes, tick marks, gridlines and secondary value axes are all honoured, and `lumMod`, `lumOff` and `satMod` colour modifiers resolve so themed charts no longer draw oversaturated. Fixes horizontal bar charts, which ignored the zero baseline and drew negative values as nothing.
- 1b6a249: Charts in a presentation render for real instead of drawing a grey placeholder. Chart parts are loaded through the slide, layout and master relationship cascade, their colours resolve against the deck theme, and the plot streams into slide primitives with an accessible label. Data labels, axis titles and per-point colours draw, and an `ofPie` group now plots as a pie rather than as columns.
- 6947366: A deck whose chart part cannot be read now opens with that chart missing, instead of the whole file being refused; slides, masters and text come through intact. Chart parts are read against a budget of their own, so a valid deck whose charts carry large cached series opens with every chart, and the part the parser declined is still written back untouched on save. That budget is one pool for the whole deck, so a chart beyond what it covers is declined the same way.
- 34541ae: PPTX decks now save with edits included, across every surface. The engine diffs the live CRDT state against a freshly seeded copy of the source package and writes back only what changed: untouched slides keep their exact source part bytes, edited slides are patched at the XML level so unmodeled markup — transitions, timing, unknown attributes — survives, and inserted or deleted slides rewrite `presentation.xml`, its relationships, and `[Content_Types].xml`. `Presentation::save()` in `betteroffice-pptx` no longer discards edits, `PresentationHandle.save()` returns the bytes on the npm core, `PptxEditor` gains a save toolbar button, Ctrl/Cmd+S, `onSave` and `fileName` props, and `save` on its `onReady` api, and the Python binding's `save`/`save_path` serialize edited decks instead of raising — `UnsupportedWriteError` is gone.

  Inside an edited paragraph, untouched runs keep their exact source markup; an edit contained in a single source run is rebuilt onto that run's properties, so hyperlinks, strikethrough, and spacing survive it. An edit spanning several source runs rewrites the span from the modeled styling — hyperlink and field bindings inside that span do not survive, which is the known write-back limitation.

- 6c7e94a: A deck snapshot persisted by the previous release opens again. Charts made the stored package a version 2 document, and the version check demanded an exact match, so every version 1 snapshot — every presentation a collaborator had already edited and saved — came back as `unsupported deck schema version` and could not be reopened.

  `open_from_update` now migrates instead of refusing. A version 1 document hydrates, its stored package is read back and rewritten in the current shape, and the document is stamped version 2, so the next snapshot the session writes is a version 2 one and the upgrade happens once. Nothing else in the document changes: the slide order, slide, shape and story containers were already identical between the two versions, and the only difference was the chart list the stored package gained. That list is optional when reading, so a package written before charts existed loads with none rather than failing on a missing field. Two clients opening the same old snapshot write the same migration and converge.

  A version this build does not know — a document from a newer release, or one whose version is missing or nonsense — is still rejected, and still reported before the stored package is parsed so the version is the error the caller sees.

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

## 0.0.2
