# @betteroffice/vsdx-react

## 0.1.0

### Minor Changes

- 6550b01: List the document's own masters as a Document Stencil. `masters()` returns every master with its name and a rendered preview display list, an inserted master instance carries its master reference through the CRDT, the snapshot and the saved XML, and its width, height and local pin stay inherited from the master.
- 87e6845: Add a View tab with grid, snap and ruler toggles. A free move snaps the dragged shape's edges and
  centre to the quarter-inch grid and to the other top-level shapes, with alignment guides while it
  holds and Alt to suppress it, and the rulers track the pointer.
- 518b018: Show where printer-paper boundaries fall across a page. The display list gains `printWidth` and `printHeight` (contract version 6), taken from the page's `PaperKind`, print orientation and margins at the page's drawing scale, and a page-break toggle on the ribbon draws the boundaries over the page. A page whose file names no printer paper reports its own extent, one sheet.
- 412ddd9: Validate a diagram against five read-only rules: dangling connectors, isolated shapes, overlapping shapes, crossing connectors and empty required shape-data rows. `validate()` and `validatePage(pageIndex)` return deterministic, ordered issues addressed by session shape id, and `IssuesPanel` renders them.

### Patch Changes

- c7072ef: Take the arrow stencil to the twenty-eight shapes Visio's Arrow Shapes stencil holds, adding the bent, U-turn, circular and arced curves alongside the striped, notched, block, quad and left-right-up blocks. Ribbon controls now read a GUARD behind a SETATREF redirect as blocked, matching what the mutation policy refuses.
- 3b1f7be: Add a second shape stencil of arrows and lines to the shapes panel. A thumbnail rail switches stencils, search filters within the active one, and a tile dragged from either stencil drops onto the canvas.
- 90dc949: Open a context menu on an empty-canvas right-click with undo, redo and add shape, drawn only from commands the ribbon registry already carries. The menu clears the selection, stays inside the viewport, closes on Escape, on an outside press, on a page switch and when the document is replaced. A hover-opened submenu no longer closes when its trigger is clicked, and command enablement now matches a GUARD call rather than any formula containing the letters GUARD.
- c5de029: Clamp the canvas backing store to the browser's maximum side length and to a per-canvas area budget, so a large page at high zoom on a high-density display no longer requests a canvas past that limit or allocates several hundred MiB for one page.
- 1bd0d84: Edit a connector's filed route. Core adds `setConnectorRoute`, which rewrites a 1-D connector's `Geometry` rows from scene-space points through the mutation policy, and the page display list now reports per-connector selection chrome with the glue kind of each endpoint. The React editor paints one handle per straight segment and drags a segment perpendicular only, re-plumbing the rest of the route.
- a841283: Drop the shape context menu when the shape it was opened for leaves the selection, so a menu opened on a deleted shape no longer reappears against the next shape selected.
- 02d3866: Resolve the `Control` section and draw its yellow diamond handles on the selected shape, positioned through the shape's own transform and its group ancestors. Dragging a handle writes the row's `X` and `Y` through the mutation policy in a single transaction, so the gesture is one undo entry and a guarded axis refuses the whole drag instead of committing half of it. Handles whose `XCon`/`YCon` behaviour is hidden are not drawn, and an axis-locked handle keeps that axis fixed. Core adds `setControlHandle`.
- 4fb0852: Add a Drawing Explorer panel: a tree of the document's pages, shapes and group children with each shape's ShapeSheet sections, rows, formulas and values. It mirrors the engine snapshot instead of keeping its own model and shares the canvas selection, and it refuses to select a shape hidden by its layer.
- 538af44: Render Visio linear fill gradients. The renderer resolves the `FillGradient` section through masters and styles, evaluates `GradientStopColor` with the theme so `THEMEVAL` stops pick up the document theme, and emits a gradient paint carrying `FillGradientAngle`. Non-linear `FillGradientDir` values, an unresolvable stop and gradients with fewer than two stops fall back to the solid fill with a fidelity diagnostic; a stop with `GradientStopColorTrans` is painted opaque and reported. The display-list contract version moves from 4 to 5 and `Paint` gains an optional `angleDeg`.
- 079738d: Show a gradient-filled shape's first stop in the fill swatch instead of falling back to black.
- 452aea5: Cut, copy, paste and duplicate a shape, and carry a group's whole child tree with it. Core adds `addShapeWithText`, `addShapeTree` and `subtreeGlue`; the React package exports `copySelection`, `pasteEntry` and `duplicateEntry` and puts the commands on the Home tab with their Ctrl shortcuts. Glue wholly inside a copied subtree is remapped onto the copy, glue crossing the copy boundary is dropped, a copy refuses unportable content with its reason, and a paste is refused when it came from another document or would leave a reference to a shape it did not copy.
- 29e488f: Add `probeCellWrites`, which asks the engine's mutation policy what a write to each cell would do without writing it, and have the editor gate its controls on that answer. The editor previously re-implemented the GUARD and SETATREF walk in TypeScript, so its idea of what was writable could drift from what the engine actually enforces; it now reads locks, GUARD interception and SETATREF redirection from the one policy that decides the write.
- c0474cc: Disable the line weight and line pattern inputs when the mutation policy would refuse the write, and
  follow SETATREF redirects when deciding that a cell is guarded. The error banner can now be dismissed
  and clears itself on the next successful edit.
- 42c4686: Render only shapes on visible layers and add a layers panel with per-layer visibility toggles. Core adds layer-aware rendering with `pageLayers`, `setLayerVisible` and `clearLayerVisibility`. Hiding a layer clears the selection of shapes it hides so no hidden shape keeps handles or commands.
- 6112a71: Drag on empty canvas to marquee-select every top-level shape the rubber band fully encloses, skipping shapes on hidden layers. Shift-drag adds to the selection, Escape cancels, and a drag that starts on a shape still moves it.
- f6bf09c: Select several shapes at once and keep a compound gesture on one undo entry. Undo, redo, Delete, Escape and the arrow keys now work while any editor chrome holds focus, Ctrl+A selects every visible shape on the page, and nudges, deletes and ribbon colour, rotate and flip commands apply to the whole selection. Core adds `moveShapes`, `deleteShapes` and `setCellFormulas`, each a single transaction that refuses the whole batch with no partial write when one shape's mutation policy refuses. Rotation now goes through a `Rotate` gesture, so `LockRotate` refuses it and the rotation grip disappears on shapes that cannot rotate.
- 970132b: Match Visio's arrow-key nudge: a plain arrow moves the selection by one ruler tick (1/16 inch) at any zoom, and Shift+arrow moves it by one screen pixel, capped at that tick so the Shift step is never the coarser one.
- 4658cfd: Handle Ctrl+S anywhere in the editor: it downloads the diagram instead of letting the browser save
  the page, including while a shape's text is being edited. Closing or reloading the tab now warns
  while there are unsaved changes or an open text draft.
- 193d6a4: Make the ribbon's line and colour controls behave: the line-weight field validates its text and commits once on Enter or blur, line pattern is a bounded picker, and the fill and line swatches read the rendered colour so a palette or theme fill no longer shows as black. Rewriting a cell with the formula it already holds no longer records an undo entry or sends a collaboration update, a non-positive `LineWeight` falls back to the default stroke width instead of an unpaintable one, and a guarded `Angle` hides the rotation grip and refuses the gesture with a receipt instead of committing on release.
- b22bbb4: Scroll and pan the diagram editor around a work surface with the page centred in it. Space- or middle-drag pans, ctrl+wheel zooms anchored under the pointer and clamped to the scrollable range, shift+wheel scrolls horizontally, and Escape during a pan cancels it without clearing the selection. Space is now the pan modifier, so it no longer starts text entry on the selected shape.
- 77e16ef: Read the `Property` section as typed Shape Data and add an editor panel that lists the selected shape's visible properties and writes edits through the mutation policy, so a guarded or locked property refuses the edit and reports it.
- 6c940fb: Add a rotate and flip submenu to the shape context menu, dispatching the ribbon's existing rotate-left, rotate-right, flip-horizontal and flip-vertical commands. A guarded Angle, FlipX or FlipY disables the matching entry, and a shape whose rotation and both flips are guarded disables the submenu trigger.
- 50d9e3d: Show a fill and line colour mini-toolbar above the shape context menu, flipping below the menu when it would leave the viewport and sharing the menu's Escape, outside-press and command-selection lifetime. A GUARD on FillForegnd or LineColor now disables that colour command everywhere it is offered, including the ribbon, instead of letting the pick reach the mutation policy and be refused. Open menus and submenus re-clamp when the window resizes.
- c1b60fb: Show a contextual Shape tab while a shape is selected, carrying that shape's fill, line and arrange
  commands, and hide the ribbon tabs that hold no commands instead of rendering a placeholder.
- f096841: Keep a refused text edit open with the typed draft intact, whether the commit was triggered by Escape or by clicking away, so a lock or guard refusal no longer discards what was typed.
- Updated dependencies [c7072ef]
- Updated dependencies [3b1f7be]
- Updated dependencies [90dc949]
- Updated dependencies [c5de029]
- Updated dependencies [1bd0d84]
- Updated dependencies [49bbdd4]
- Updated dependencies [02d3866]
- Updated dependencies [6550b01]
- Updated dependencies [4fb0852]
- Updated dependencies [538af44]
- Updated dependencies [3dd1830]
- Updated dependencies [452aea5]
- Updated dependencies [5c63e9d]
- Updated dependencies [29e488f]
- Updated dependencies [42c4686]
- Updated dependencies [89be4ec]
- Updated dependencies [f6bf09c]
- Updated dependencies [518b018]
- Updated dependencies [193d6a4]
- Updated dependencies [bbdd080]
- Updated dependencies [77e16ef]
- Updated dependencies [50d9e3d]
- Updated dependencies [14b1e8b]
- Updated dependencies [c1b60fb]
- Updated dependencies [657ad4f]
- Updated dependencies [dc76ac5]
- Updated dependencies [412ddd9]
  - @betteroffice/vsdx-i18n@0.1.0
  - @betteroffice/vsdx@0.1.0

## 0.0.3

### Patch Changes

- 16d33a7: Fix locale declarations for TypeScript consumers with `skipLibCheck: false` and update React editors to depend on the corrected i18n packages.
- Updated dependencies [16d33a7]
  - @betteroffice/vsdx-i18n@0.0.3
  - @betteroffice/vsdx@0.0.3

## 0.0.2

### Patch Changes

- 0df1345: Make the editor canvas focusable and give it a keyboard layer: undo and redo, Delete or Backspace,
  arrow-key nudge with a larger Shift step, and Escape to cancel a gesture or clear the selection.
  Keys are ignored while focus is in a text input, and a handle resize of a locked shape is refused
  with a translated message.
- 6fd0030: Stop the editor erroring on diagrams whose text runs carry no diagnostics: the renderer omits the empty `diagnostics` field, so the type is now optional and the reader guards it.
- fa8e835: Replace the desktop-style ribbon with Visio for the web's flat command bar: a single 45px row of
  icon buttons separated by thin rules, with no group labels, under a tab strip whose active tab
  carries a text-width underline. Tabs with no commands say so instead of rendering blank.
- d242ac2: Track the pointer while dragging or Shift-drag resizing a shape, painting a live outline of the
  landing position on the editor's overlay canvas instead of moving the shape only on release.

  Add `modelPointToCanvas`, the exact forward of `canvasPointToModel`, and commit a gesture only once
  it passes the same drag threshold the preview uses, so the preview and the committed geometry can
  never disagree.

- 3c20209: Add a right-click menu for the selected shape with keyboard navigation and z-order submenus that
  stay inside the viewport, draw the rotation grip as a circle on a stalk, and give the shapes panel
  a header, a search box and a category rail. A queued drag preview follows the latest Shift state.
- ee33b67: Draw a shape's geometry when a paint channel cannot be resolved: the fill or stroke falls back to the file's default foreground and the display list carries a diagnostic naming what failed.
  The editor reports those diagnostics alongside its text notices.
- da68f90: Drag a shape from the gallery onto the canvas and it lands where it was dropped, already selected.
  A tile click still inserts at the page centre and cascades there, per page, so repeats no longer
  stack. Every master inserts at its own aspect ratio, one inch tall, and a right-click that misses
  every shape now drops the selection along with the menu it closes.
- 7d549fb: Add a selection frame with resize handles and a rotation grip. Commit handle resizes atomically and evaluate formula LocPins at the new size to keep the preview and opposite edge in place.
- eda21ff: Add in-place shape text editing to the VSDX editor. Text edits run through the mutation policy with a typed receipt, are undoable and authorized on remote updates, and saving patches only the edited shape's `Text` element.
- 1391c9b: Expand the standard shape gallery to thirty-two shapes, including a real ellipse at a 3:2 aspect
  ratio. Every gallery preview is derived from the geometry the insert writes, so a shape lands in
  the proportions it was picked in.
- Updated dependencies [0df1345]
- Updated dependencies [6fd0030]
- Updated dependencies [fa8e835]
- Updated dependencies [d242ac2]
- Updated dependencies [3c20209]
- Updated dependencies [5408e87]
- Updated dependencies [ee33b67]
- Updated dependencies [78e3184]
- Updated dependencies [7d549fb]
- Updated dependencies [eda21ff]
- Updated dependencies [1391c9b]
  - @betteroffice/vsdx-i18n@0.0.2
  - @betteroffice/vsdx@0.0.2
