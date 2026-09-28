# @betteroffice/pptx

## 0.2.0

### Minor Changes

- 80341ac: Expose opt-in operation profiles for XLSX edits and display lists, and PPTX layout, edits, and history operations.
- ad64aa4: **Breaking:** collaboration updates now carry deck schema 2.2, and a 0.1.x update is migrated to 2.2 when it is opened. 0.1.x clients cannot open a document this version saved or migrated, so mixed-version collaboration across this upgrade is not supported: upgrade every client that shares a stored update.
- 7665432: Add version-checked, all-or-nothing PPTX edit batches. `PresentationHandle.version()`, `readContent()` and `findText()` return slides and story text with the session version they were read at; `validateEdits()` and `applyEdits()` resolve every step against that version and either commit text insertion, replacement, deletion and formatting, paragraph alignment, speaker notes, and shape rectangle, fill and outline changes as one transaction and one undo step, or return a typed refusal (`stale-version`, `missing-target`, `ambiguous-target`, `content-mismatch`, `overlapping-steps`, `unsupported`, `invalid-step`, `limit-exceeded`) with the deck untouched. `history: "none"` keeps a batch out of undo history, and `source` records provenance only. `PptxEditorApi` gains the same operations, flushing pending input first, refreshing the editor once after an applied batch, and refusing with `read-only` while the editor is read-only. The Rust `DeckSession` and `Presentation` and the Python `Presentation` expose the same API, and proposal previews and acceptance now share the batch staging. Ids and versions are session-scoped. The Rust `EditStep` and `EditFailureCode` enums are non-exhaustive.
- affeb36: Expose the PPTX editor's commands as `PptxEditorApi.commands`: serializable state with a stated reason for every disabled command, descriptors with platform-aware shortcuts, and `execute` that runs after pending input such as a decoding picture and checks availability again. Hosts compose built-in controls with their own actions through the `toolbar` and `showToolbar` props, `PptxCommandProvider`, `usePptxCommand`/`usePptxCommandState`, `EditorToolbar mode="commands"`, `ToolbarCommand`, `ToolbarCommandButton`, `ToolbarCommandSelect` and `ToolbarOverflow`, inside or outside the editor. The prop-based `EditorToolbar`, `Toolbar` and `useEditorToolbar` keep working and are deprecated. Canvas and panel proposal review, keyboard shortcuts and the default toolbar run through the same commands; a picture lands on the slide it was chosen for. Narrow toolbars move groups into a keyboard-accessible More menu, and `ToolbarDropdown` follows the menu or dialog pattern of its content. New locale keys cover command labels and disabled reasons. The command and toolbar composition API is experimental and may change in minor releases. `@betteroffice/pptx` adds `anchorCaret` and `resolveCaretAnchor`: plain-data caret positions that follow later edits, undo, redo and remote updates.
- 2eb2c55: Expose host save interception, awaited input flushing, and pointer position queries. Add manual undo capture, explicit boundaries, and undoable comment repositioning. `PptxEditorApi.save()` now throws while input or a pointer gesture is pending (await `flushPendingInput()` first) and after the presentation is replaced, instead of returning bytes that miss accepted input.
- 911a294: Insert a picture onto a slide from the editor. The image mints its own media part, content-type default and relationship on save; `PptxEditor` gains a small "Insert image" icon button next to the text-box tool, and `PresentationHandle` gains `addPicture`. `addPicture` takes PNG, JPEG, GIF, BMP, TIFF, WebP and SVG images up to 8 MiB, checked before the picture reaches the deck, keeping oversized bytes out of collaboration updates.
- 911a294: Reorder a shape's paint order on its slide: bring to front, send to back, and step it forward or backward. `PresentationHandle` gains `bringShapeToFront`, `sendShapeToBack`, `bringShapeForward` and `sendShapeBackward`, and `PptxEditor`'s shape-formatting toolbar gains an "Arrange" menu for them.
- 73d405c: Export PPTX as read-only structured content and Markdown. `PresentationHandle.exportStructured()` and `exportMarkdown()` read the committed deck with the version they were read at and change nothing; `exportPptxStructured(bytes)`, `exportPptxMarkdown(bytes)` and `renderPptxMarkdown(content)` do the same headless as snapshots. Schema version 1 JSON lists slides in deck order with their original index, the current shape tree depth first (a documented reading order, not one inferred from geometry), text stories whose paragraphs carry level, effective alignment, authored `bulletJson`, the list marker resolved through the layout, master and text-style cascade the renderer uses, and runs with formatting marks, links, soft line breaks, cached field results and zero-width placeholders for inline content such as equations; tables keep their grid, spans and merge continuations with each cell's current text; pictures, video, audio, charts, SmartArt, embedded objects and shape-tree elements the deck model does not hold become placeholders with alternative text, the relationships they reference, and the parts and external targets those resolve to. Every record has a deterministic id and an anchor (`slide`, `shape`, `range`, `notes`, `comment`, or `sourcePart`), a `range` anchor being the batch text target in `readContent()` offsets so a host can pass it straight into `applyEdits()`, and records read from the file carry their source part, SHA-256, element path, `sldId` and `cNvPr` id. Hidden slides and shapes (unrepresented elements included), plain-text speaker notes and comments are opt-in, notes-page text outside the notes is diagnosed even when the notes are empty, each slide exports its own shapes (an empty placeholder as empty, layout and master shapes as an `inherited-content-omitted` diagnostic), and every omission is a diagnostic. Slides an older collaboration update does not record the visibility of are exported with `hidden: null` and a `visibility-unknown` diagnostic. `maxBlocks` and `maxBytes` are checked before each slide, shape, story, table row and comment is read, stop the export at a whole record, keep truncated tables renderable, and set `truncated`; unusable limits are refused as `invalid-options` or `limit-exceeded` with `target: null`. Markdown renders slide headings, title placeholders as headings, literal list markers, pipe or entity-escaped HTML tables, object placeholders, notes and comments, with a `<!-- pptx-export:N -->` marker per block mapped to its anchor; text, alt text and metadata are escaped and kept on one line where a line break would open a block; links are kept only when the scheme of the whole target, with character references and percent escapes decoded and whitespace and controls removed, is http, https, mailto, tel, ftp or absent; and `renderPptxMarkdown` refuses content that breaks its contract (anchor kinds and owners, UTF-16 ranges, paragraphs that tile their story unless the export was truncated, unique ids, cells in column order, merge origins, nesting, one formatting mark of each kind per run) as `invalid-content`, bounds the records and metadata it validates, and stops building a block once it cannot fit `maxBytes`. The Rust `DeckSession` and `Presentation`, `betteroffice_pptx::export_pptx_structured`, and the Python `Presentation.export_structured()` and `export_pptx_structured()` expose the same export. The parser now records hidden slides (`p:sld/@show`), and the paragraph cascade and list numbering the renderer uses live in `pptx_edit::paragraph`.
- 8a5e6b8: Render PPTX pictures stored as SVG in both backends. The browser packages decode them with the browser's own SVG support, and the Rust and Python `render_png` paths rasterize them natively: shapes, paths, fills, strokes, gradients, clip paths, `<use>` and class-based stylesheets draw, under a sandbox that budgets the cost of expanding and painting a document before any of it runs. A tiled SVG repeats at its intrinsic size.
- ab3d722: Add a format-owned text search API to presentation handles.

### Patch Changes

- 6c30f4a: Draw charts closer to Office: an overlaid legend sits on the plot and a side legend is laid out from its font, number formats apply their sections, stacked category labels break onto lines, series fills keep their opacity, a value-axis title stands on its side and its ticks start at the minimum, line points sit in the middle of their categories, unstyled markers take Office's size and symbol, a gridline whose `spPr` draws no line is hidden, line and scatter labels go where `dLblPos` says, chart text is centred where the plot geometry centres it, and a lone bar series varies its colours unless `varyColors` is `0`.
- 9639a6b: Charts whose cached data names a very large point index no longer allocate a slot for every index up to it, which let a small chart use hundreds of megabytes and run WebAssembly out of memory. A cache now spans no more than its `c:ptCount`, points past it are ignored, and empty slots between points count against the chart's point limit; missing points still draw as gaps.
- 6963a67: Honour `a:normAutofit` as PowerPoint renders it: the stored `fontScale` (0.1–1.0) scales sizes, rounded to whole points below 1.0, and `lnSpcReduction` (0–0.9) is subtracted from percentage line spacing, including the implicit single-spaced default, so shrink-to-fit bodies keep PowerPoint's font size and line pitch.
- 3fb2bf7: Number `a:buAutoNum` paragraphs as one list when they repeat the same `startAt`: PowerPoint writes a list's start on every one of its paragraphs, so a four-item list marked `startAt="4"` now draws 4, 5, 6, 7 and one marked `startAt="1"` draws a), b), c). A paragraph that declares a different start still opens a new list.
- c02a145: Paint slide background pictures and honour `p:sld/@showMasterSp` on the layout. A `p:bg` declaring a `a:blipFill`, or a `p:bgRef` resolving to one through the theme's `a:bgFillStyleLst`, now paints as a full-slide image instead of a flat grey, and `p:bgRef` resolves the referenced fill style rather than only its colour override. A slide that turns off master shapes also drops the layout's own decoration, which is what PowerPoint draws, since the master reaches the slide through the layout.
- 030505a: pptx: resolve PowerPoint's built-in table styles, and stop dressing a table in
  the one `tableStyles.xml` only nominates. `ppt/tableStyles.xml` carries the
  styles a deck edited, so a table naming a style PowerPoint has never had to
  write out — `{5C22544A…}` "Medium Style 2 - Accent 1" above all, the default
  every new table and every python-pptx table takes — found nothing and rendered
  as bare text on the slide background: no header fill, no banding, no borders,
  no white bold header. 17 of the 59 tables in the 103-deck corpus, across 11
  decks, name a style their own package never defines. Three of those styles are
  now resolved from a built-in catalogue, transcribed from the definitions
  PowerPoint itself serialises when a deck does edit them. Separately, the
  `def` attribute names the style the authoring UI hands a _new_ table, not a
  fallback for one that names none, so a table with an empty `a:tblPr` no longer
  picks it up: `pptarena-010` was painting ten converted forms in `accent1`
  tint 20% where PowerPoint leaves them white.

  On the 103-deck fidelity corpus (101 scored, 1,033 pages) the mean rises
  0.89784 to 0.89900, +0.00116; 9 decks improve, led by `pptarena-061` +0.04292,
  `pptarena-071` +0.03352 and `pptarena-025` +0.02781, with no deck regressing,
  92 of 101 decks byte-identical, and page counts exact in both arms.

- 8e8f97a: Render EMF pictures that wrap a bitmap: `presentationImageBlob` unwraps an enhanced metafile whose only ink is one unscaled `EMR_STRETCHDIBITS` covering its bounds into the BMP it carries, alongside the bitmap-only WMF wrappers already handled. The blit must be an unrotated, uncropped `SRCCOPY` of a `BI_RGB` DIB that fills the metafile frame.
- 030505a: pptx: an empty paragraph now takes the height its own `a:endParaRPr` asks for. Empty paragraphs are how authors write vertical spacers, and PowerPoint sizes each one from the run properties it carries; the renderer read the size the shape's list style would have given a run instead, so on `pptarena-042` a 40pt spacer was drawn at the 211.2pt title default — 5.28 times too tall — and the error accumulated down the body until text left its placeholder and crossed the footer. The property was parsed and never read. It is now resolved through the same slide/layout/master cascade the rest of the paragraph properties already use, so it reaches text that arrives as a collaborative story as well as text read straight from the package, and a paragraph without one keeps the size it inherits today. On the 103-deck fidelity corpus (101 scored, 1,033 pages) the mean rises 0.89198 to 0.89765, +0.00568; 42 decks improve, led by `pptarena-052` +0.1101, `pptarena-055` +0.0736 and `pptarena-042` +0.0625, and one deck moves -0.000042 while its text lands 4px closer to PowerPoint.
- 3e0c311: Measure the `hexagon`, `parallelogram`, `trapezoid` and `octagon` adjust against the shortest side, and pin it at the spec's aspect-scaled maximum, so wide shapes no longer draw their slant or corner at a fraction of the width. The trapezoid defaults to the spec's 25000, and the hexagon honours its `vf` height factor, pinned so the corners stay on the frame.
- 5114756: Drag and resize placeholders that take their geometry from the layout or master, from the frame they are drawn in, and draw them with the rotation and flips they inherit. The snapshot carries that geometry as `inherited`; the first `moveShape`, `resizeShape` or `setShapeRect` makes the whole transform the placeholder's own, so it keeps its size and orientation live and after a save. The Rust crates add `Placeholder::matches`, the placeholder matching that rendering and editing share, and the facade re-exports `InheritedGeometry`.
- 030505a: pptx: a picture stored as an EMF metafile now draws when the file sets a clip. The metafile player treated any unlisted record as fatal and abandoned the whole drawing, so a single `EXTSELECTCLIPRGN` — record 12 of a typical file, and a no-op where every corpus occurrence resets to the default region — silently blanked the picture. Rectangular clips are now tracked on the device context, survive `SAVEDC`/`RESTOREDC`, and narrow the shape they draw into. 18 of 69 corpus metafiles decode where 10 did before, and on the 103-deck fidelity corpus `pptarena-042` moves 0.7229 to 0.8451 with no deck regressing.
- 6963a67: Open `spcBef` and `spcAft` percentages over a single line rather than the bare text size, the base PowerPoint measures them against, so the stock Office master's 20% gap sits where PowerPoint puts it between bulleted paragraphs.
- 58f9bfb: Render arc, cube, leftBrace, rightBrace, wedgeRectCallout, ribbon2, swooshArrow, and circularArrow with their PowerPoint geometry. Preserve separate fills, shaded faces, and open outlines; complete cloud callout details and folded-corner shading, use circular rounded-rectangle corners, and honor stroke joins in browser and native rendering.

  Expose `geometryFallback` diagnostics on shapes, picture fills and masks.

- 60c79dd: Draw eleven more PowerPoint preset shapes. `donut`, `noSmoking`, `corner`, `foldedCorner`, `mathMultiply`, `bentArrow`, `ribbon`, `ellipseRibbon`, `cloudCallout`, `wedgeEllipseCallout` and `wedgeRoundRectCallout` fell back to a plain rectangle; they now follow their ECMA-376 definitions, including the elliptical arcs `arcTo` measures by polar angle rather than by ellipse parameter.
- 7faab61: Free a proposal's cached preview once it is rejected or accepted, so a long review session no longer grows memory with every proposal, and previewing a rejected or accepted proposal fails with "no proposal" again instead of returning its stale preview.
- ad64aa4: A collaboration update saved by 0.1.x and reopened with its source file saves an unedited deck byte-identically again: run caps and colour-mapped `tx1`/`bg1` run colours are restored from the source for the text the 0.1.x seed wrote, in long and edited stories too, and the preset adjustments the 0.1.x seed stored as defaults, where no collaborator has changed them, move to the current defaults instead of drawing default stars as near-polygons. In JavaScript, a matching source that cannot be reattached now throws instead of opening a session that cannot save, and a save refuses rather than drop run baseline or spacing it could not recover.
- 7f158c7: Support `a:rPr/@cap`, so a run that asks for all caps or small caps is drawn that way. `all` uppercases the run for drawing and `small` also draws the lowercase stretches at Word's 0.8× small-cap size; `none` turns off an inherited setting. The value cascades from a slide master's and layout's `a:defRPr` the way the other run properties do, and the uppercasing itself is now shared with DOCX rather than reimplemented. Casing is a display property: the stored run text keeps the author's casing, so the editor's story, the caret offsets it reports and the saved parts are all unchanged. Measured against the PPTArena corpus tail, where `cap="all"` reaches titles in `pptarena-051`, `pptarena-052`, `pptarena-054` and body runs in `pptarena-034`.
- 3830d79: A character a face lacks is drawn from a script fallback face registered with `registerFallbackFont`.
- 6963a67: Measure a single-spaced PowerPoint line as 1.2 em of the line's largest font, the pitch PowerPoint 16.113 uses for every face, so percentage line spacing no longer inherits the substituted face's own ascent, descent and line gap and multi-line bodies stop drifting away from PowerPoint down the shape. Super- and subscript ink still pushes the line box out past that pitch.
- 27bf1fc: Resolve `tx1`/`bg1`/`tx2`/`bg2` scheme colours through the slide's colour map: the master's `p:clrMap` and any `p:clrMapOvr/a:overrideClrMapping` on its layout or the slide itself now decide which `a:clrScheme` slot each name reaches, so dark-master decks paint their backgrounds, shape fills and text the way PowerPoint does. One shared resolver feeds the render, snapshot and save projections, so they cannot disagree about a slot.
- e0d12f3: Render a slide at the page extent PowerPoint exports. A slide's page box is a whole number of points, so `p:sldSz` snaps there before the pixel scale, and the canvas backing store now covers a fractional extent by rounding up, matching what the raster path already does. An A4 deck whose `p:sldSz` is 10691813×7559675 EMU renders 1755×1240 px at 150 DPI, the size of PowerPoint's own PDF page, instead of coming up a pixel short.
- 18e1f32: Draw star presets with the spec's inner radius, per-preset default adjustment, and frame-filling radius factors.
- d76b4db: Resolve pictures stored only as SVG: a blip with no `r:embed` of its own takes its media from the Office SVG extension (`asvg:svgBlip`), for pictures and shape picture fills alike. A blip that carries both keeps its raster image.
- c5f1467: Draw the shape a Wingdings, Wingdings 2, Wingdings 3 or Webdings `a:buChar` addresses. Those faces reach their glyphs by font position, through a private-use cmap at `U+F0xx` or the raw byte, so the character a deck stores for one of them names a slot rather than the character to draw. Slots for circles, squares, hollow boxes, diamonds, and the solid triangles and arrowheads now resolve to the nearest Unicode character the bundled faces cover. The shapes come from the faces' own glyph names and the sizes from matching each glyph's ink width against the candidates', confirmed against PowerPoint's own render of the corpus decks. An `a:buChar` under a text face, or one that is already a real Unicode character, is drawn exactly as authored.
- 6c30f4a: Render slide text closer to PowerPoint: autofit keeps point spacing and rounds sizes to whole points below a scale of 1.0, automatic numbering skips blank paragraphs, text is centred within its line spacing, named families use their own text metrics and heavy weights, text outside placeholders takes the presentation's `defaultTextStyle`, and right-to-left paragraphs are laid out from the right.
- 030505a: pptx: a tab in slide text now advances to a tab stop. Text layout had no notion
  of tabs at all, so a `U+0009` went to the shaper like any other character and
  came back as the fallback face's `.notdef` — a box where that face draws one,
  nothing where it does not, and in both cases an advance that owed nothing to the
  stops the file declares. Tabbed columns collapsed into the run before them, and
  the wrong width moved the wrap point and grew the table row around it.
  `a:pPr/@defTabSz` and `a:pPr/a:tabLst` are now read and inherited through the
  same slide/layout/master cascade the other paragraph properties use, and a tab
  takes the first declared stop past the pen or, failing that, the next multiple
  of the default pitch — one inch when nothing declares one — measured from the
  text area's left edge, painting no glyph and never reaching past the line. A
  hanging indent adds the implicit stop at the paragraph margin that its first tab
  lands on; a marker already owns that space here, so only an unmarked hanging
  indent has any. 20 of the 103 corpus decks carry tab characters, 17 of them on
  slides. On the 103-deck fidelity corpus (101 scored, 1,033 pages) the mean rises
  0.89765 to 0.89784, +0.00019; 16 decks improve, led by `pptarena-033` +0.00695,
  `pptarena-050` +0.00325, `pptarena-043` +0.00282 and `pptarena-010` +0.00181,
  with no deck regressing and page counts exact in both arms.
- af6292e: pptx: a slide whose background fill is fully transparent now renders on white paper instead of leaving the page unpainted. A master declaring `solidFill` at `alpha="0"` produced a paint rather than no paint, so the existing white fallback never fired and the slide came out as a hole onto whatever sat behind the canvas — which a PNG export then flattened to black. PowerPoint paints those slides white.
- 1f84618: Render slides that carry a picture the browser cannot decode: `paintSlide` treats a rejected `resolveImage` as an unresolved picture, so the remaining primitives on that slide, and every later slide, still paint. A resolver that rejects — `createImageBitmap` raises `InvalidStateError: The source image could not be decoded` for EMF, vector WMF and other media no browser decodes — previously rejected the whole `paintSlide` call, blanking the slide. The picture's own outline still strokes, matching an unresolved asset.

## 0.1.1

## 0.1.0

### Minor Changes

- d6ba9da: Add session-local PPTX agent proposals with atomic acceptance, stale-target checks, rendered previews, and one-step Undo. Expose the workflow in Rust, TypeScript, and Python. Show inline text diffs and previous/proposed shape bounds on the React slide canvas, with proposal selection, accept/reject controls, and a review panel for before/after previews.
- d6e6e91: Read, add, reply to, resolve and remove PowerPoint comments in both the legacy and modern formats, saved by patching the existing comment XML in place (deck schema 2.1; older clients reject new updates until upgraded).
- bf84789: Set paragraph alignment from the presentation toolbar.
- 387f239: Add PPTX-to-PNG export for Rust, Python, and browsers, plus an Export PNG button
  in the React editor.

  Rust: enable the `raster` feature; exhaustive `Error` matches must handle the new
  `Raster` variant.

- 21b48f3: Edit per-slide speaker notes that persist through saves and collaboration updates, and present slides fullscreen from the React editor with keyboard navigation.
- ef5cdee: Render DrawingML tables with styled cells, merged regions, borders, and text-driven row heights.

### Patch Changes

- 93971b5: Remove outdated early-release warnings from package READMEs and link the JavaScript guide and changelogs.
- bd69e9e: Draw DOCX plus presets as twelve-vertex crosses from the ECMA-376 arm guide instead of rectangular fallback, with default and missing-adjustment shapes sharing the same path.
- d926fb0: Correct tint proportions in shared OOXML color resolution so lighter theme colors blend toward white as authored.
- cae162d: Drop redundant buffer copies around the wasm boundary and per collaboration update.
- 6ae0b92: Read the shapes an `mc:AlternateContent` contributes to a shape tree — the first `mc:Choice` whose `Requires` namespaces are implemented, otherwise the `mc:Fallback` — and keep edits, deletions and no-edit saves aligned with the branch that holds them.
- 010865c: Measure block arrow heads from the shortest side while preserving shaft widths across aspect ratios.
- 2877aba: Render automatic list numbers with bullet formatting, preserve explicit restarts when editing and saving, and migrate collaboration snapshots to deck schema 2.1.
- 1f30ea0: Apply picture duotone, biLevel, grayscale and colour-change effects in Canvas, PNG exports and the native viewer, preserving alpha and migrating older collaboration snapshots.
- abb1e2c: Apply a picture's `a:lum` brightness and contrast in Canvas, PNG exports and the native viewer.
- 899aac5: Round an automatic value axis to whole `{1, 2, 5} x 10^k` steps and widen its unpinned ends to the next step, so a stacked bar no longer ends on the plot edge.
- c4985a8: Respect explicit chart data label settings that disable every field.
- bfc3231: Reserve space for top and bottom chart legends, wrap their entries to fit, and center PowerPoint chart titles.
- 89a2134: Stroke a chart series' line at the width its own `c:ser/c:spPr/a:ln` declares instead of a fixed 2px, and draw no line at all when that outline is `a:noFill`.

  Migrate collaboration snapshots to deck schema 2.1, importing series lines from a reattached source.

- d2aaf9c: Paint a chart's own `c:chartSpace` fill instead of a white ground, and stroke each axis line with its own `a:ln` colour and width, or not at all under `a:noFill`.

  Migrate collaboration snapshots to deck schema 2.1, importing chart-space fills and axis lines from a reattached source.

- c9b72bf: Draw chart titles, axis labels, legends and data labels in the family, slant and character spacing their `c:txPr` declares, instead of the theme minor font upright and untracked. A chart title's own `c:rich` run properties now override the paragraph default they sit under, and a reattached source refreshes the text properties of a stored chart. A DOCX chart paints the tracking its text properties declare instead of only reserving room for it.
- bc34dfc: Parse connector shapes and preserve legacy collaboration updates when editing and saving.
- 0c9b52e: Render numeric custom geometry paths, including elliptical arcs and separate path fills and strokes. Preserve custom geometry in deck schema 2.1 collaboration snapshots.
- 69167fe: Paint justified lines at their caret positions and keep editor gestures consistent.
- 413499c: Preserve weight and slant when substituting missing presentation fonts, and choose
  the nearest fallback style consistently regardless of face registration order.
- 2044df7: Render gradient outlines across browser, raster, and native backends while preserving their paint through theme inheritance, width edits, and legacy snapshot migration.
- 8b48e8d: Render unordered gradient stops correctly in slide display lists and raster output while preserving equal-position stop order and source XML.
- 54fdaa0: Skip hidden slide shapes and hidden groups' descendants when painting and hit-testing.

  Deck schema 2.1 migrates existing 1.0 and 2.0 documents by recovering hidden flags from stored package data. Older clients reject the new schema.

  `ShapeSnapshot.hidden` is optional and omitted when false, preserving unchanged snapshot JSON. Only hidden shapes store a Yrs key; the schema stamp changes for all decks.

- cca2618: Transpose horizontal bar chart axes, preserve category direction, and reserve space for category labels, axis titles, and secondary value ticks.
- 2b639b9: Render triangle, open arrow, stealth, diamond, and oval line ends on the PPTX
  canvas, preserving their independent width and length settings.
- 2c90c17: Apply inherited list styles and render character bullets with their own formatting while preserving caret positions and migrating collaboration snapshots to deck schema 2.1.
- a61781d: Replay EMF GRADIENTFILL records as shaded bands and the BITBLT raster operations that carry no source bitmap, instead of rejecting the whole metafile.
- 25c7ea3: Render supported EMF and WMF vector pictures with crops, masks and fill rules, and preserve OLE fallback pictures through deck schema 2.1 collaboration updates.
- d280c87: Keep hyperlinks, fields, and every other unmodeled run markup when a text edit spans several source runs. Surviving text is written back onto the run it came from, typed text extends the run before it, a field whose text changed becomes a plain run so PowerPoint does not overwrite it, and a run whose text is deleted takes its markup with it.
- 22ce4e9: Parse `a:effectLst/a:outerShdw` and paint the drop shadow of filled and outlined shapes in Canvas and PNG exports.

  Preserve shadow scale and alignment across deck schema 2.1 snapshots. Bound per-slide shadow work in Canvas and PNG exports.

- 3d95068: Apply inherited paragraph line spacing, preserve baselines for expanded point spacing, and migrate collaboration snapshots to deck schema 2.1.
- 088d177: Narrow a paragraph's wrap width by its right margin, inherited from the master text styles, a layout or shape `a:lstStyle`, or the paragraph itself, and recover it from an attached source for collaboration snapshots written before deck schema 2.1.
- 875d556: Open the space a paragraph asks for with `a:spcBef` and `a:spcAft`, inherited from the master text styles, a layout or shape `a:lstStyle`, or the paragraph itself, and recover them from an attached source for collaboration snapshots written before deck schema 2.1.
- 70e7394: Crop pictures to their `srcRect` and clip and outline their preset masks in Canvas, PNG exports, and the native viewer. Preserve JSON for uncropped rectangular pictures.
- acab663: Paint a picture's outer shadow from the silhouette of its alpha in Canvas and PNG exports, and keep a picture-filled shape's shadow through the image rewrite.
- 0824bff: Scale chevron and homePlate adjustments from the shortest side, allow their points to span the full width, and normalize DOCX preset guide values consistently.
- 069e4d6: Render superscript and subscript runs at their shifted baselines and preserve their formatting through edits and saved decks.
- 1e86217: Render gradient-filled text using its lowest valid stop and preserve authored gradients across text edits until their color changes.
- 89f8f7b: Preserve each text run's colour, weight, italic, underline, and size while keeping identically styled adjacent runs grouped.
- f5d9fd9: Render and preserve run character spacing through edits, collaboration snapshots, and saved presentations.
- e5c4521: Collapse the never-released deck schema chain into a single 2.1 step, so a 1.0 or 2.0 snapshot still migrates and reopens exactly as before.
- 5c015e9: Refuse shape adjustment edits on custom geometry instead of silently dropping them on save, and report an exhausted shape id space instead of overflowing it.
- 9274a2b: Render stretched picture fills through shape geometry and retain their source data across collaboration snapshots.

  Migrate collaboration snapshots to deck schema 2.1, preserving source imports and edited text.

- 25d4ee4: Use shape font-reference colours above master text defaults while preserving run, paragraph, and placeholder colours.
- 7fdc0ee: Evaluate slide-number fields on masters and layouts, counting from the presentation's first slide number.

  Collaboration snapshots use deck schema 2.1. Older snapshots migrate to deck
  schema 2.1; missing starting numbers default to one. Readers supporting only
  older schemas reject 2.1 snapshots.

- 60113a3: Parse `a:tbl` into a real table model: the column grid, row heights, cell spans and merge continuations, direct cell fills and borders, and the `a:tblPr` style flags. A cell's `a:tcPr` anchoring, text direction and margins fold into its text body.

  Recover a table's geometry and cell formatting from a reattached source, folded into the unreleased schema 2.1.

- 253d680: Add the `table` display-list primitive, a container whose cells paint clipped to the table rectangle in every backend, and raise the display-list contract version to 2.
- a139ae9: Stop a table's first cell painting as loose slide text on top of the graphic frame that holds it.
- 051830e: Parse `ppt/tableStyles.xml` and resolve a cell through the table style cascade: `wholeTbl`, the row band, `firstCol`/`lastCol`, `firstRow`/`lastRow`, then the cell's own `a:tcPr`, with each part's borders applied against the sides of its own region and an `a:noFill` edge clearing what a lower part set. A reattached source restores the styles a stored package never carried.
- 07d72ce: Keep shape text unmirrored and honor vertical text direction, insets, and caret positions without changing ordinary horizontal rotations.
- 1af946f: Render overflowing text at its intended size and anchor across backends, preserve explicit clipping, and keep transformed text clickable while migrating collaboration snapshots to schema 16.
- a3b2acd: Resolve DrawingML font references to the theme's major or minor script face, using its Latin face when the requested script slot is empty.
- 2710a41: Resolve PPTX theme fill and line references, including background fills and placeholder colour transforms. Preserve explicit shape and placeholder formatting. Preserve font reference colours from the existing text-style resolver.

  Migrate deck snapshots to deck schema 2.1, preserving edits, numbering and source ordinals.

- b1f5c91: Render embedded TIFF pictures in browser presentations by converting them to PNG inside the PPTX WASM boundary. Uncompressed, LZW, PackBits and deflate sources are supported, including grayscale, RGB, palette and CMYK images; other compressions are skipped.
- bbd80c5: Turn `eaVert` and `mongolianVert` text boxes the way `vert` turns, run Mongolian columns left to right, and stack `wordArtVert` and `wordArtVertRtl` one character to a line.
- 7ce54d6: Render bitmap-only WMF images in the slide editor and expose presentationImageBlob for canvas image resolvers. Preserve original media bytes when saving.
- 863b70e: Reject a WMF picture whose MOVETO or LINETO record is too short to hold its point, instead of drawing it without that line.

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
