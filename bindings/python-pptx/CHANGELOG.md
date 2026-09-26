# @betteroffice/python-pptx

## 0.2.0

### Minor Changes

- 8a5e6b8: Render PPTX pictures stored as SVG in both backends. The browser packages decode them with the browser's own SVG support, and the Rust and Python `render_png` paths rasterize them natively: shapes, paths, fills, strokes, gradients, clip paths, `<use>` and class-based stylesheets draw, under a sandbox that budgets the cost of expanding and painting a document before any of it runs. A tiled SVG repeats at its intrinsic size.
- 67e0ae4: Release the accumulated DOCX, PPTX, and XLSX engine improvements in the Python bindings as minor updates.

### Patch Changes

- 6963a67: Honour `a:normAutofit` the way PowerPoint renders it: the stored `fontScale` is applied verbatim and `lnSpcReduction` is subtracted from percentage line spacing — including the implicit single-spaced default — so shrink-to-fit bodies keep PowerPoint's font size and line pitch instead of being re-fitted at render time.
- 3fb2bf7: Number `a:buAutoNum` paragraphs as one list when they repeat the same `startAt`: PowerPoint writes a list's start on every one of its paragraphs, so a four-item list marked `startAt="4"` now draws 4, 5, 6, 7 and one marked `startAt="1"` draws a), b), c). A paragraph that declares a different start still opens a new list.
- 5114756: Drag and resize placeholders that take their geometry from the layout or master, from the frame they are drawn in, and draw them with the rotation and flips they inherit. The snapshot carries that geometry as `inherited`; the first `moveShape`, `resizeShape` or `setShapeRect` makes the whole transform the placeholder's own, so it keeps its size and orientation live and after a save. The Rust crates add `Placeholder::matches`, the placeholder matching that rendering and editing share, and the facade re-exports `InheritedGeometry`.
- 6963a67: Open `spcBef` and `spcAft` percentages over a single line rather than the bare text size, the base PowerPoint measures them against, so the stock Office master's 20% gap sits where PowerPoint puts it between bulleted paragraphs.
- 58f9bfb: Render arc, cube, leftBrace, rightBrace, wedgeRectCallout, ribbon2, swooshArrow, and circularArrow with their PowerPoint geometry. Preserve separate fills, shaded faces, and open outlines; complete cloud callout details and folded-corner shading, use circular rounded-rectangle corners, and honor stroke joins in browser and native rendering.

  Expose `geometryFallback: true` when unsupported geometry is replaced by a rectangle, including picture fills and masks.

- 7f158c7: Support `a:rPr/@cap`, so a run that asks for all caps or small caps is drawn that way. `all` uppercases the run for drawing and `small` also draws the lowercase stretches at Word's 0.8× small-cap size; `none` turns off an inherited setting. The value cascades from a slide master's and layout's `a:defRPr` the way the other run properties do, and the uppercasing itself is now shared with DOCX rather than reimplemented. Casing is a display property: the stored run text keeps the author's casing, so the editor's story, the caret offsets it reports and the saved package are all unchanged — an untouched deck still saves byte-identically. Measured against the PPTArena corpus tail, where `cap="all"` reaches titles in `pptarena-051`, `pptarena-052`, `pptarena-054` and body runs in `pptarena-034`.
- 6963a67: Measure a single-spaced PowerPoint line as 1.2 em of the line's largest font, the pitch PowerPoint 16.113 uses for every face, so percentage line spacing no longer inherits the substituted face's own ascent, descent and line gap and multi-line bodies stop drifting away from PowerPoint down the shape. Super- and subscript ink still pushes the line box out past that pitch.
- 27bf1fc: Resolve `tx1`/`bg1`/`tx2`/`bg2` scheme colours through the slide's colour map: the master's `p:clrMap` and any `p:clrMapOvr/a:overrideClrMapping` on its layout or the slide itself now decide which `a:clrScheme` slot each name reaches, so dark-master decks paint their backgrounds, shape fills and text the way PowerPoint does. One shared resolver feeds the render, snapshot and save projections, so they cannot disagree about a slot.
- d76b4db: Resolve pictures stored only as SVG: a blip with no `r:embed` of its own takes its media from the Office SVG extension (`asvg:svgBlip`), for pictures and shape picture fills alike. A blip that carries both keeps its raster image.
- c5f1467: Draw the shape a Wingdings, Wingdings 2, Wingdings 3 or Webdings `a:buChar` addresses. Those faces reach their glyphs by font position, through a private-use cmap at `U+F0xx` or the raw byte, so the character a deck stores for one of them names a slot rather than the character to draw. Each slot now resolves to the nearest Unicode character the bundled faces cover — circles, squares, hollow boxes, diamonds, and the solid triangles and arrowheads — with a plain bullet standing in for a slot that has no covered equivalent. The shapes come from the faces' own glyph names and the sizes from matching each glyph's ink width against the candidates', confirmed against PowerPoint's own render of the corpus decks. An `a:buChar` under a text face, or one that is already a real Unicode character, is drawn exactly as authored.

## 0.1.0

### Minor Changes

- d6ba9da: Add session-local PPTX agent proposals with atomic acceptance, stale-target checks, rendered previews, and one-step Undo. Expose the workflow in Rust, TypeScript, and Python. Show inline text diffs and previous/proposed shape bounds on the React slide canvas, with proposal selection, accept/reject controls, and a review panel for before/after previews.
- 387f239: Add PPTX-to-PNG export for Rust, Python, and browsers, plus an Export PNG button
  in the React editor.

  Rust: enable the `raster` feature; exhaustive `Error` matches must handle the new
  `Raster` variant.

### Patch Changes

- 4bcc61f: Release the GIL during python-pptx open, render, save, font registration and update application, without copying the input bytes to do it.

## 0.0.2

### Patch Changes

- 34541ae: PPTX decks now save with edits included, across every surface. The engine diffs the live CRDT state against a freshly seeded copy of the source package and writes back only what changed: untouched slides keep their exact source part bytes, edited slides are patched at the XML level so unmodeled markup — transitions, timing, unknown attributes — survives, and inserted or deleted slides rewrite `presentation.xml`, its relationships, and `[Content_Types].xml`. `Presentation::save()` in `betteroffice-pptx` no longer discards edits, `PresentationHandle.save()` returns the bytes on the npm core, `PptxEditor` gains a save toolbar button, Ctrl/Cmd+S, `onSave` and `fileName` props, and `save` on its `onReady` api, and the Python binding's `save`/`save_path` serialize edited decks instead of raising — `UnsupportedWriteError` is gone.

  Inside an edited paragraph, untouched runs keep their exact source markup; an edit contained in a single source run is rebuilt onto that run's properties, so hyperlinks, strikethrough, and spacing survive it. An edit spanning several source runs rewrites the span from the modeled styling — hyperlink and field bindings inside that span do not survive, which is the known write-back limitation.

- 12d83a9: `betteroffice-pptx` publishes to PyPI. The distribution joins the release train's publish matrix, and a `workflow_dispatch` workflow can publish one binding on its own — to create a PyPI project from a pending Trusted Publisher, or to fill in a wheel a release run dropped. Publishing is Trusted Publishing only; the workflow refuses to run while a repository-scoped `PYPI_API_TOKEN` exists.
- 98b9225: Both distributions point their PyPI `Documentation` link at
  https://docs.betteroffice.dev/docs/python, a page that names the package, gives
  the install line and a first example. The link used to land on the documentation
  root, which never mentioned Python at all.
- a52867e: `betteroffice-pptx` is a new PyPI package that puts the PPTX engine behind a typed Python API, next to `betteroffice-xlsx`. `Presentation.open` parses a deck into the same resident CRDT model the browser editor uses, and `snapshot`, `slide`, `story` and `media` read it back as plain Python objects — shapes with resolved fill and outline colours, geometry in EMU, text grouped into stories, paragraphs and runs, and embedded images as bytes you can write straight to disk. `insert_slide`, `move_slide`, `delete_slide`, `add_text_box`, `add_shape`, `move_shape`, `resize_shape`, `set_shape_fill`, `set_shape_stroke`, `set_shape_adjust`, `remove_shape`, `insert_text`, `delete_text`, `format_text` and `insert_paragraph_break` each return a receipt naming what changed, including the rect before and after a transform and the adjustment values after the engine clamped them. `register_font` and `render_slide` produce the renderer's display list as JSON, so a server can lay a slide out with the same contract the canvas paints. `open_collaborative`, `state_vector`, `state_as_update`, `diff` and `apply_update` are the same Yrs primitives the xlsx binding exposes, under the same names, so a project using both packages sees one contract.

  `save` and `save_path` serialize the deck with every accepted edit applied: untouched slides keep their exact source part bytes, edited slides are patched at the XML level, and the read-only `is_edited` reports whether the engine has accepted an edit since the deck was opened — only an edit the engine accepted sets it; an edit that raised leaves the flag untouched. An unsupported preset geometry, an unknown adjustment guide, an out-of-range client ID or an unknown parse limit raises `ValueError` rather than being ignored, `delete_text` raises `RangeError` on a range that crosses a paragraph boundary rather than swallowing the break — `format_text` styles across paragraphs instead, as one undoable edit — and an oversized collaboration payload is refused before it is copied. Rendering refuses when no font has been registered, because nothing is embedded in this wheel.

  Only a deck from `open_collaborative` is a replica. A deck from `open` or `open_path` has no client ID of its own, so two of them would author under the same identity and never converge; `state_vector`, `state_as_update`, `diff` and `apply_update` therefore raise `NotCollaborativeError` on one, and `is_collaborative` reports which kind a caller holds. `betteroffice-xlsx` refuses the same four calls, but it refuses them in the engine, on `WorkbookMode::Standalone`; this guard lives in the Python binding alone, so the Rust `betteroffice-pptx` facade still hands those primitives to a standalone deck. Pushing the check down into the engine is follow-up work.

  A `Presentation` is pinned to the thread that opened it _and must be released there too_: the engine's undo manager is not `Send`, so unlike `betteroffice-xlsx` this binding is `unsendable`. Touching a deck from another thread raises `pyo3_runtime.PanicException`, which is a `BaseException` rather than an `Exception` and so escapes `except Exception`. Dropping one on another thread raises nothing at all — pyo3 skips the Rust destructor, writes an unraisable `RuntimeError`, and strands roughly 1.5 MB per deck. That applies to the cyclic garbage collector as well, which frees a cycle wherever it happens to run, so giving each worker its own deck is not on its own enough; the deck has to become garbage on its owning thread. Engine calls hold the GIL for their duration, though the file I/O releases it: `open_path`'s read, and the writes in `save_path`, `Media.write` and `DisplayList.write`.

- bdedb87: Both Python READMEs now describe what the bindings actually do. The
  `betteroffice-xlsx` page said `save` regenerated the package and dropped parts
  the model does not cover; saving has preserved charts, drawings, pivot tables,
  comments, macros and custom XML since part preservation landed, so the Status
  section now states the preservation and the limits that remain — an edited sheet
  is still reserialized from the model, and this binding exposes no structural
  edits at all. The `betteroffice-pptx` page said unregistered families fell back
  to a metrics-only path; `render_slide` raises `RenderError: no font has been
registered for slide text` instead, so the layout section leads with that error
  and then shows registering a face, and says what a family you did not register
  resolves to once one exists.

  Both READMEs name the import next to the install line, because `pip install
betteroffice-xlsx` gives `import betteroffice_xlsx`, and both name the incumbent
  they are compared against in the opening paragraph rather than a hundred lines
  down. The xlsx API table stops presenting the static `Workbook.open_collaborative`
  as an instance method, and gains `value`, `formula`, `proposals`,
  `merged_ranges`, `last_calculation`, `sheet_index`, `can_undo`/`can_redo`, and
  the `now_serial` keyword that supplies the clock `TODAY()` and `NOW()` read.
  `StaleProposalError` now documents its way out, `accept_proposal(id,
force=True)`. The proposals example prints the value it actually produces, and
  the pptx snippets no longer assume the first slide has a shape or that the shape
  bears text.

  Both distributions declare `Operating System :: OS Independent` and per-minor
  `Programming Language :: Python` classifiers for 3.9 through 3.13, so PyPI's
  version filter finds them, swap the `openpyxl-alternative` and
  `python-pptx-alternative` keywords nobody searches for the bare project names,
  and add a `Changelog` project URL.
