# @betteroffice/pptx-i18n

## 0.3.0

## 0.2.0

### Minor Changes

- affeb36: Expose the PPTX editor's commands as `PptxEditorApi.commands`: serializable state with a stated reason for every disabled command, descriptors with platform-aware shortcuts, and `execute` that runs after pending input such as a decoding picture and checks availability again. Hosts compose built-in controls with their own actions through the `toolbar` and `showToolbar` props, `PptxCommandProvider`, `usePptxCommand`/`usePptxCommandState`, `EditorToolbar mode="commands"`, `ToolbarCommand`, `ToolbarCommandButton`, `ToolbarCommandSelect` and `ToolbarOverflow`, inside or outside the editor. The prop-based `EditorToolbar`, `Toolbar` and `useEditorToolbar` keep working and are deprecated. Canvas and panel proposal review, keyboard shortcuts and the default toolbar run through the same commands; a picture lands on the slide it was chosen for. Narrow toolbars move groups into a keyboard-accessible More menu, and `ToolbarDropdown` follows the menu or dialog pattern of its content. New locale keys cover command labels and disabled reasons. The command and toolbar composition API is experimental and may change in minor releases. `@betteroffice/pptx` adds `anchorCaret` and `resolveCaretAnchor`: plain-data caret positions that follow later edits, undo, redo and remote updates.
- 911a294: Insert a picture onto a slide from the editor. The image mints its own media part, content-type default and relationship on save; `PptxEditor` gains a small "Insert image" icon button next to the text-box tool, and `PresentationHandle` gains `addPicture`. `addPicture` takes PNG, JPEG, GIF, BMP, TIFF, WebP and SVG images up to 8 MiB, checked before the picture reaches the deck, keeping oversized bytes out of collaboration updates.
- d11dff3: Host editor plugins in `PptxEditor` through the new `plugins`, `pluginGrants` and `onPluginError` props. The plugin API is experimental and may change in minor releases. A plugin created with `definePptxPlugin` can contribute a docked panel (left, right or bottom of the slide), an overlay on the slide canvas, and commands registered as `plugin:<pluginId>/<id>`, whose results carry the plugin's own failure codes or a refused edit batch unchanged, with toolbar entries (`PptxPluginToolbar` places them in replacement chrome) and shortcuts; `ToolbarCommandButton`, `ToolbarCommand`, `usePptxCommand` and `usePptxCommandState` accept those ids. Each activation receives `load`, `document-change`, `selection-change`, `mode-change`, `layout-change` and `grants-change` events, restricted clients (versioned `readContent`, `findText` and `validateEdits`, commands, granted edit batches, and `goToSlide`, `selectShape` and `selectText` navigation that keeps focus unless asked), slide geometry that converts slide EMU or pixels into overlay pixels and reports rendered shape bounds, version-guarded state and `onCleanup` disposers. Plugins read, validate and navigate by default; built-in commands and edit batches need an explicit grant, rechecked with `readOnly` immediately before every write, and mutating built-in commands refuse plugins with `unsupported-policy`. Every contribution renders behind its own error boundary with a plugin-scoped command context, and a failing plugin is stopped and reported without affecting the editor or other plugins. A superseded slide paint no longer draws over the slide painted after it. New locale keys cover plugin panels, the plugin toolbar group and the new disabled reasons.
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
