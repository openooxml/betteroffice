# @betteroffice/vsdx

## 0.1.0

### Minor Changes

- 49bbdd4: Add `addConnectedShape` to the VSDX diagram handle: it inserts a shape and a
  connector glued to an existing shape in one transaction and one undo entry, and
  refuses the edit when the glue does not resolve to a connection point.
- 6550b01: List the document's own masters as a Document Stencil. `masters()` returns every master with its name and a rendered preview display list, an inserted master instance carries its master reference through the CRDT, the snapshot and the saved XML, and its width, height and local pin stay inherited from the master.
- 538af44: Render Visio linear fill gradients. The renderer resolves the `FillGradient` section through masters and styles, evaluates `GradientStopColor` with the theme so `THEMEVAL` stops pick up the document theme, and emits a gradient paint carrying `FillGradientAngle`. Non-linear `FillGradientDir` values, an unresolvable stop and gradients with fewer than two stops fall back to the solid fill with a fidelity diagnostic; a stop with `GradientStopColorTrans` is painted opaque and reported. The display-list contract version moves from 4 to 5 and `Paint` gains an optional `angleDeg`.
- 3dd1830: Add `addFreeConnector` to the VSDX diagram handle: it inserts a connector glued
  at its begin only, leaving the end where the draft puts it, and refuses the edit
  when the glue does not resolve to a connection point.
- 518b018: Show where printer-paper boundaries fall across a page. The display list gains `printWidth` and `printHeight` (contract version 6), taken from the page's `PaperKind`, print orientation and margins at the page's drawing scale, and a page-break toggle on the ribbon draws the boundaries over the page. A page whose file names no printer paper reports its own extent, one sheet.
- 14b1e8b: Render shape shadows. A shape's `ShdwPattern`, offsets and colour resolve through its own cells, its cached values and the theme effect scheme, ride in the display list as a `shadow` field, and paint on the canvas in device pixels. The display-list contract version moves from 6 to 7; a consumer pinning 6 must update.
- 657ad4f: Export a diagram as SVG or PNG. `exportSvg()` returns one vector page per diagram page, sized from the PageSheet, with text kept as `<text>`; `exportPng(pageIndex, scale)` rasterizes one page at `scale` times the 96 dpi display list. Both walk the same ordered display list as the canvas and the PDF exporter, and both paint shape shadows.
- 412ddd9: Validate a diagram against five read-only rules: dangling connectors, isolated shapes, overlapping shapes, crossing connectors and empty required shape-data rows. `validate()` and `validatePage(pageIndex)` return deterministic, ordered issues addressed by session shape id, and `IssuesPanel` renders them.

### Patch Changes

- c5de029: Clamp the canvas backing store to the browser's maximum side length and to a per-canvas area budget, so a large page at high zoom on a high-density display no longer requests a canvas past that limit or allocates several hundred MiB for one page.
- 1bd0d84: Edit a connector's filed route. Core adds `setConnectorRoute`, which rewrites a 1-D connector's `Geometry` rows from scene-space points through the mutation policy, and the page display list now reports per-connector selection chrome with the glue kind of each endpoint. The React editor paints one handle per straight segment and drags a segment perpendicular only, re-plumbing the rest of the route.
- 02d3866: Resolve the `Control` section and draw its yellow diamond handles on the selected shape, positioned through the shape's own transform and its group ancestors. Dragging a handle writes the row's `X` and `Y` through the mutation policy in a single transaction, so the gesture is one undo entry and a guarded axis refuses the whole drag instead of committing half of it. Handles whose `XCon`/`YCon` behaviour is hidden are not drawn, and an axis-locked handle keeps that axis fixed. Core adds `setControlHandle`.
- 452aea5: Cut, copy, paste and duplicate a shape, and carry a group's whole child tree with it. Core adds `addShapeWithText`, `addShapeTree` and `subtreeGlue`; the React package exports `copySelection`, `pasteEntry` and `duplicateEntry` and puts the commands on the Home tab with their Ctrl shortcuts. Glue wholly inside a copied subtree is remapped onto the copy, glue crossing the copy boundary is dropped, a copy refuses unportable content with its reason, and a paste is refused when it came from another document or would leave a reference to a shape it did not copy.
- 5c63e9d: Keep a group's transform on its display-list node instead of baking it into every child, and derive that transform from the group's own cells rather than the live bounding box of its children. Editing or removing one child no longer moves or resizes its siblings, hit testing and line jumps compose ancestor transforms, and `bounds_affine` drops its child-extent argument.
- 29e488f: Add `probeCellWrites`, which asks the engine's mutation policy what a write to each cell would do without writing it, and have the editor gate its controls on that answer. The editor previously re-implemented the GUARD and SETATREF walk in TypeScript, so its idea of what was writable could drift from what the engine actually enforces; it now reads locks, GUARD interception and SETATREF redirection from the one policy that decides the write.
- 42c4686: Render only shapes on visible layers and add a layers panel with per-layer visibility toggles. Core adds layer-aware rendering with `pageLayers`, `setLayerVisible` and `clearLayerVisibility`. Hiding a layer clears the selection of shapes it hides so no hidden shape keeps handles or commands.
- 89be4ec: Migrate VSDX diagrams stored under the previous CRDT schema. Stories written before in-place text editing held resolved text tokens as JSON under the same schema version as the plain text that replaced them, so reopening such a document showed and saved a JSON blob as every shape's text. The schema version now distinguishes the two and a stored document is carried forward when it is opened.
- f6bf09c: Select several shapes at once and keep a compound gesture on one undo entry. Undo, redo, Delete, Escape and the arrow keys now work while any editor chrome holds focus, Ctrl+A selects every visible shape on the page, and nudges, deletes and ribbon colour, rotate and flip commands apply to the whole selection. Core adds `moveShapes`, `deleteShapes` and `setCellFormulas`, each a single transaction that refuses the whole batch with no partial write when one shape's mutation policy refuses. Rotation now goes through a `Rotate` gesture, so `LockRotate` refuses it and the rotation grip disappears on shapes that cannot rotate.
- 193d6a4: Make the ribbon's line and colour controls behave: the line-weight field validates its text and commits once on Enter or blur, line pattern is a bounded picker, and the fill and line swatches read the rendered colour so a palette or theme fill no longer shows as black. Rewriting a cell with the formula it already holds no longer records an undo entry or sends a collaboration update, a non-positive `LineWeight` falls back to the default stroke width instead of an unpaintable one, and a guarded `Angle` hides the rotation grip and refuses the gesture with a receipt instead of committing on release.
- bbdd080: Add `setShapeData`, a single engine op that writes a batch of shape-data row values for one shape as one undo entry. Every row is decided before anything is written, so a refusal anywhere leaves the document untouched, and the caller gets one receipt per row saying whether it was written and why not. Date, duration and currency rows refuse rather than being rewritten as text, a number row refuses a non-numeric value, and an empty formula is rejected.
- 77e16ef: Read the `Property` section as typed Shape Data and add an editor panel that lists the selected shape's visible properties and writes edits through the mutation policy, so a guarded or locked property refuses the edit and reports it.
- dc76ac5: Keep a shape's text upright when the shape is flipped. `FlipX` or `FlipY` mirrored the glyphs with the outline, and a chain that flipped both axes painted them upside down; the text box now reflects about the shape's own box once per flipped axis, so the label stays unmirrored over the geometry the flip moved, inside flipped groups too.

## 0.0.3

## 0.0.2

### Patch Changes

- 6fd0030: Stop the editor erroring on diagrams whose text runs carry no diagnostics: the renderer omits the empty `diagnostics` field, so the type is now optional and the reader guards it.
- d242ac2: Track the pointer while dragging or Shift-drag resizing a shape, painting a live outline of the
  landing position on the editor's overlay canvas instead of moving the shape only on release.

  Add `modelPointToCanvas`, the exact forward of `canvasPointToModel`, and commit a gesture only once
  it passes the same drag threshold the preview uses, so the preview and the committed geometry can
  never disagree.

- 5408e87: Add an engine API for creating glued VSDX connectors in one edit transaction. Preserve glue through collaboration and saving, and omit records whose shapes have been deleted.
- ee33b67: Draw a shape's geometry when a paint channel cannot be resolved: the fill or stroke falls back to the file's default foreground and the display list carries a diagnostic naming what failed.
  The editor reports those diagnostics alongside its text notices.
- 78e3184: Honour Geometry section `NoFill`, `NoLine` and `NoShow` controls when rendering shapes and connectors.
  Render visible geometry with only the paint channels it uses, while retaining diagnostics for unsupported controls.
- 7d549fb: Add a selection frame with resize handles and a rotation grip. Commit handle resizes atomically and evaluate formula LocPins at the new size to keep the preview and opposite edge in place.
- eda21ff: Add in-place shape text editing to the VSDX editor. Text edits run through the mutation policy with a typed receipt, are undoable and authorized on remote updates, and saving patches only the edited shape's `Text` element.
