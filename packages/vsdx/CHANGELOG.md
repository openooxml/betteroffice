# @betteroffice/vsdx

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
