# @betteroffice/vsdx-react

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
