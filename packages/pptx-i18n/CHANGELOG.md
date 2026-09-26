# @betteroffice/pptx-i18n

## 0.2.0

### Minor Changes

- 911a294: Insert a picture onto a slide from the editor. The image mints its own media part, content-type default and relationship on save; `PptxEditor` gains a small "Insert image" icon button next to the text-box tool, and `PresentationHandle` gains `addPicture`. Unsupported MIME types and images over 8 MiB are rejected before the picture reaches the deck, keeping oversized bytes out of collaboration updates.
- 911a294: Reorder a shape's paint order on its slide: bring to front, send to back, and step it forward or backward. `PresentationHandle` gains `bringShapeToFront`, `sendShapeToBack`, `bringShapeForward` and `sendShapeBackward`, and `PptxEditor`'s shape-formatting toolbar gains an "Arrange" menu for them.

## 0.1.1

### Patch Changes

- 16d33a7: Fix locale declarations for TypeScript consumers with `skipLibCheck: false` and update React editors to depend on the corrected i18n packages.

## 0.1.0

### Minor Changes

- d6ba9da: Add session-local PPTX agent proposals with atomic acceptance, stale-target checks, rendered previews, and one-step Undo. Expose the workflow in Rust, TypeScript, and Python. Show inline text diffs and previous/proposed shape bounds on the React slide canvas, with proposal selection, accept/reject controls, and a review panel for before/after previews.

### Patch Changes

- 93971b5: Remove outdated early-release warnings from package READMEs and link the JavaScript guide and changelogs.
- 915dbaa: Translate the 0.1.0 toolbar, presentation, speaker-notes, and agent-proposal strings for every shipped locale so non-English users no longer see English fallback on the new controls.

## 0.0.4

### Patch Changes

- 34541ae: PPTX decks now save with edits included, across every surface. The engine diffs the live CRDT state against a freshly seeded copy of the source package and writes back only what changed: untouched slides keep their exact source part bytes, edited slides are patched at the XML level so unmodeled markup — transitions, timing, unknown attributes — survives, and inserted or deleted slides rewrite `presentation.xml`, its relationships, and `[Content_Types].xml`. `Presentation::save()` in `betteroffice-pptx` no longer discards edits, `PresentationHandle.save()` returns the bytes on the npm core, `PptxEditor` gains a save toolbar button, Ctrl/Cmd+S, `onSave` and `fileName` props, and `save` on its `onReady` api, and the Python binding's `save`/`save_path` serialize edited decks instead of raising — `UnsupportedWriteError` is gone.

  Inside an edited paragraph, untouched runs keep their exact source markup; an edit contained in a single source run is rebuilt onto that run's properties, so hyperlinks, strikethrough, and spacing survive it. An edit spanning several source runs rewrites the span from the modeled styling — hyperlink and field bindings inside that span do not survive, which is the known write-back limitation.

## 0.0.3

### Patch Changes

- 5212690: Google Slides-style editor toolbar for the PPTX editor: new-slide split button
  with layout picker, undo/redo, zoom, select and text-box tools, and contextual
  text formatting that also applies to whole shapes on selection. Text formatting
  now spans paragraph boundaries as a single undoable operation, double/triple
  click select word/paragraph, and roundRect corners render circular per the
  OOXML adj value instead of stretching with the shape.
- b87185f: Shape insertion and styling: a Slides-style shape picker inserts preset
  geometries (rectangles, ellipse, polygons, stars, arrows, chevron) by click
  or drag, and selected shapes get contextual fill, border color, border width,
  and corner-radius controls backed by new undoable, collaboration-native
  addShape/setShapeFill/setShapeStroke/setShapeAdjust engine operations.

## 0.0.2
