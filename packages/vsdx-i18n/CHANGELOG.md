# @betteroffice/vsdx-i18n

## 0.1.0

### Patch Changes

- c7072ef: Take the arrow stencil to the twenty-eight shapes Visio's Arrow Shapes stencil holds, adding the bent, U-turn, circular and arced curves alongside the striped, notched, block, quad and left-right-up blocks. Ribbon controls now read a GUARD behind a SETATREF redirect as blocked, matching what the mutation policy refuses.
- 3b1f7be: Add a second shape stencil of arrows and lines to the shapes panel. A thumbnail rail switches stencils, search filters within the active one, and a tile dragged from either stencil drops onto the canvas.
- 90dc949: Open a context menu on an empty-canvas right-click with undo, redo and add shape, drawn only from commands the ribbon registry already carries. The menu clears the selection, stays inside the viewport, closes on Escape, on an outside press, on a page switch and when the document is replaced. A hover-opened submenu no longer closes when its trigger is clicked, and command enablement now matches a GUARD call rather than any formula containing the letters GUARD.
- 6550b01: List the document's own masters as a Document Stencil. `masters()` returns every master with its name and a rendered preview display list, an inserted master instance carries its master reference through the CRDT, the snapshot and the saved XML, and its width, height and local pin stay inherited from the master.
- 4fb0852: Add a Drawing Explorer panel: a tree of the document's pages, shapes and group children with each shape's ShapeSheet sections, rows, formulas and values. It mirrors the engine snapshot instead of keeping its own model and shares the canvas selection, and it refuses to select a shape hidden by its layer.
- 452aea5: Cut, copy, paste and duplicate a shape, and carry a group's whole child tree with it. Core adds `addShapeWithText`, `addShapeTree` and `subtreeGlue`; the React package exports `copySelection`, `pasteEntry` and `duplicateEntry` and puts the commands on the Home tab with their Ctrl shortcuts. Glue wholly inside a copied subtree is remapped onto the copy, glue crossing the copy boundary is dropped, a copy refuses unportable content with its reason, and a paste is refused when it came from another document or would leave a reference to a shape it did not copy.
- 42c4686: Render only shapes on visible layers and add a layers panel with per-layer visibility toggles. Core adds layer-aware rendering with `pageLayers`, `setLayerVisible` and `clearLayerVisibility`. Hiding a layer clears the selection of shapes it hides so no hidden shape keeps handles or commands.
- f6bf09c: Select several shapes at once and keep a compound gesture on one undo entry. Undo, redo, Delete, Escape and the arrow keys now work while any editor chrome holds focus, Ctrl+A selects every visible shape on the page, and nudges, deletes and ribbon colour, rotate and flip commands apply to the whole selection. Core adds `moveShapes`, `deleteShapes` and `setCellFormulas`, each a single transaction that refuses the whole batch with no partial write when one shape's mutation policy refuses. Rotation now goes through a `Rotate` gesture, so `LockRotate` refuses it and the rotation grip disappears on shapes that cannot rotate.
- 518b018: Show where printer-paper boundaries fall across a page. The display list gains `printWidth` and `printHeight` (contract version 6), taken from the page's `PaperKind`, print orientation and margins at the page's drawing scale, and a page-break toggle on the ribbon draws the boundaries over the page. A page whose file names no printer paper reports its own extent, one sheet.
- 193d6a4: Make the ribbon's line and colour controls behave: the line-weight field validates its text and commits once on Enter or blur, line pattern is a bounded picker, and the fill and line swatches read the rendered colour so a palette or theme fill no longer shows as black. Rewriting a cell with the formula it already holds no longer records an undo entry or sends a collaboration update, a non-positive `LineWeight` falls back to the default stroke width instead of an unpaintable one, and a guarded `Angle` hides the rotation grip and refuses the gesture with a receipt instead of committing on release.
- 77e16ef: Read the `Property` section as typed Shape Data and add an editor panel that lists the selected shape's visible properties and writes edits through the mutation policy, so a guarded or locked property refuses the edit and reports it.
- 50d9e3d: Show a fill and line colour mini-toolbar above the shape context menu, flipping below the menu when it would leave the viewport and sharing the menu's Escape, outside-press and command-selection lifetime. A GUARD on FillForegnd or LineColor now disables that colour command everywhere it is offered, including the ribbon, instead of letting the pick reach the mutation policy and be refused. Open menus and submenus re-clamp when the window resizes.
- c1b60fb: Show a contextual Shape tab while a shape is selected, carrying that shape's fill, line and arrange
  commands, and hide the ribbon tabs that hold no commands instead of rendering a placeholder.
- 412ddd9: Validate a diagram against five read-only rules: dangling connectors, isolated shapes, overlapping shapes, crossing connectors and empty required shape-data rows. `validate()` and `validatePage(pageIndex)` return deterministic, ordered issues addressed by session shape id, and `IssuesPanel` renders them.

## 0.0.3

### Patch Changes

- 16d33a7: Fix locale declarations for TypeScript consumers with `skipLibCheck: false` and update React editors to depend on the corrected i18n packages.

## 0.0.2

### Patch Changes

- 0df1345: Make the editor canvas focusable and give it a keyboard layer: undo and redo, Delete or Backspace,
  arrow-key nudge with a larger Shift step, and Escape to cancel a gesture or clear the selection.
  Keys are ignored while focus is in a text input, and a handle resize of a locked shape is refused
  with a translated message.
- fa8e835: Replace the desktop-style ribbon with Visio for the web's flat command bar: a single 45px row of
  icon buttons separated by thin rules, with no group labels, under a tab strip whose active tab
  carries a text-width underline. Tabs with no commands say so instead of rendering blank.
- 3c20209: Add a right-click menu for the selected shape with keyboard navigation and z-order submenus that
  stay inside the viewport, draw the rotation grip as a circle on a stalk, and give the shapes panel
  a header, a search box and a category rail. A queued drag preview follows the latest Shift state.
- eda21ff: Add in-place shape text editing to the VSDX editor. Text edits run through the mutation policy with a typed receipt, are undoable and authorized on remote updates, and saving patches only the edited shape's `Text` element.
- 1391c9b: Expand the standard shape gallery to thirty-two shapes, including a real ellipse at a 3:2 aspect
  ratio. Every gallery preview is derived from the geometry the insert writes, so a shape lands in
  the proportions it was picked in.
