# @betteroffice/docx-react

## 0.3.0

### Minor Changes

- 4884924: Add `onSaveRequest` for host-controlled File > Save and Cmd/Ctrl+S, plus awaited `flushPendingInput()` on the editor refs. Built-in saves flush pending input before serialization, and concurrent UI save requests share one workflow.

### Patch Changes

- Updated dependencies [8eddb2e]
- Updated dependencies [ca68499]
- Updated dependencies [6db9411]
- Updated dependencies [69b7cca]
- Updated dependencies [c087612]
- Updated dependencies [59d71c4]
- Updated dependencies [0e9ca88]
- Updated dependencies [9b2fe7b]
- Updated dependencies [ebacc4e]
- Updated dependencies [9b2fe7b]
- Updated dependencies [72aed1b]
- Updated dependencies [b1d9284]
- Updated dependencies [023fad0]
- Updated dependencies [e4cb228]
- Updated dependencies [8273450]
- Updated dependencies [b46ad04]
- Updated dependencies [4eb1d90]
- Updated dependencies [9b2fe7b]
- Updated dependencies [e151d79]
- Updated dependencies [e5ad702]
- Updated dependencies [2958935]
- Updated dependencies [41f508b]
- Updated dependencies [1247270]
  - @betteroffice/docx@0.3.0
  - @betteroffice/docx-i18n@0.3.0

## 0.2.1

### Patch Changes

- 16d33a7: Fix locale declarations for TypeScript consumers with `skipLibCheck: false` and update React editors to depend on the corrected i18n packages.
- Updated dependencies [16d33a7]
  - @betteroffice/docx-i18n@0.2.1
  - @betteroffice/docx@0.2.1

## 0.2.0

### Minor Changes

- 295f42f: Keep one local undo history across document stories, group rapid keystrokes in WebAssembly, and preserve native undo in other inputs. Replace story-scoped history helpers with session-wide tracking and changed-story reporting.

  Migrate each removed API as follows: `historyStory()` returns the changed stories via `historyStories()` (sorted, empty before the first local edit instead of `null`); `undoDepth()` and `redoDepth()` are gone, query `canUndo()` and `canRedo()` instead; `markUndoGroup(startDepth)` is gone, rapid keystrokes now coalesce in WebAssembly with no host bookkeeping; `applyLocalUpdate(update, story)` drops its story argument and becomes `applyLocalUpdate(update)`; `beginUndoCapture(story, includeTableStories?)` drops its arguments and becomes `beginUndoCapture()`; `computeLayout()` no longer returns `blocks` and `measures`, read them lazily from `getLayoutKernelInputs(computation.layout)` as `measured` and `options`.

### Patch Changes

- 5069ad2: Keep accepted spreadsheet proposals undoable in collaborative sessions, preserve pending proposals through remote edits, and require a refreshed review when calculated previews change. Reject document suggestions that overlap partially tracked text. Existing public signatures and wire fields remain unchanged.
- 93971b5: Remove outdated early-release warnings from package READMEs and link the JavaScript guide and changelogs.
- 0b0a90a: Reuse successfully painted pages while scrolling and release temporary canvas buffers after replay.
- 43fad65: Reduce large-document interaction cost with per-line selection bands, lazy Unicode caret stops, compact retained-page shift replay, revision-bound lazy measured inputs, and stable page rendering identities.
- 1efec27: Improve DOCX fidelity with Word's 10 pt fallback for undeclared font sizes, short-paragraph widow control, and corrected table padding, minimum row heights, repeated headers, and rotated image sizing. Align automatic and wrapped tabs to the page grid, keep wrapped text metrics on their own lines, and hide list markers on page-break-only paragraphs. Preserve authored formatting and document state.
- 1d830df: Add a CDN-only font provider, settle Japanese font preflight without retry loops, and preserve floating header shapes without inflating body margins. Load and save alternate main-document filenames through their package relationships, and forward layout failures through the editor error callback.
- b1f5c91: Render embedded TIFF pictures in browser documents by converting them to PNG inside the DOCX parse WASM boundary. Uncompressed, LZW, PackBits and deflate sources are supported, including grayscale, RGB, palette and CMYK images; other compressions are skipped.
- 846b5d6: Fall back to the main-thread engine when the resident worker crashes, times out, or answers corruptly mid-input so typing survives worker failures. Replay the pending keystroke on the main-thread engine and keep genuine engine-level input rejections surfacing as errors.
- 0019657: Recover from resident worker crashes, WebAssembly traps, and unanswered requests so the editor can fall back to the main-thread engine. Reset retained worker frames and queries when switching engines so fresh main-thread frames render immediately.
- Updated dependencies [93971b5]
- Updated dependencies [b351bbe]
- Updated dependencies [2061849]
- Updated dependencies [4511b9a]
- Updated dependencies [4bf205b]
- Updated dependencies [0664bd3]
- Updated dependencies [5c04bc4]
- Updated dependencies [2c56acd]
- Updated dependencies [d4f4b85]
- Updated dependencies [a958376]
- Updated dependencies [cb2c6fe]
- Updated dependencies [356f67a]
- Updated dependencies [d4f4b85]
- Updated dependencies [2c658b6]
- Updated dependencies [297f43d]
- Updated dependencies [43fad65]
- Updated dependencies [6393137]
- Updated dependencies [4cc8dc0]
- Updated dependencies [1efec27]
- Updated dependencies [9281f7e]
- Updated dependencies [6779738]
- Updated dependencies [c1f9684]
- Updated dependencies [4d27f1e]
- Updated dependencies [b2ca63b]
- Updated dependencies [bd69e9e]
- Updated dependencies [20e913c]
- Updated dependencies [1d830df]
- Updated dependencies [6ce2439]
- Updated dependencies [e42ff8d]
- Updated dependencies [036f83e]
- Updated dependencies [a117530]
- Updated dependencies [9e4656f]
- Updated dependencies [1dc0e41]
- Updated dependencies [295f42f]
- Updated dependencies [56c3ca4]
- Updated dependencies [6f0e36d]
- Updated dependencies [284f0c4]
- Updated dependencies [d4f4b85]
- Updated dependencies [8adcd04]
- Updated dependencies [d4f4b85]
- Updated dependencies [b1f5c91]
- Updated dependencies [2ee434c]
- Updated dependencies [1a5ef23]
- Updated dependencies [1d0f41d]
- Updated dependencies [846b5d6]
- Updated dependencies [0019657]
- Updated dependencies [73cea54]
- Updated dependencies [d926fb0]
- Updated dependencies [c9b72bf]
- Updated dependencies [1e46a6f]
  - @betteroffice/docx@0.2.0
  - @betteroffice/docx-i18n@0.2.0

## 0.1.0

### Minor Changes

- 9540e23: Clicking a footnote or an endnote now opens it for editing. The caret lands under the pointer, typing goes into that note, and what is typed reaches the saved file. Escape leaves the note; clicking back in the body leaves it and places the body caret under the pointer.

  Undo follows the part being edited: the first edit in a newly opened part replaces the undo scope and discards the previous part's history.

  A note area is document text, not chrome, so a single click opens it — where a header or footer band, which repeats on every page and sits in the margin, still needs a double-click. The click that opens the note also places its caret: the position it resolved belongs to a story the editor is not on yet, so it waits for that story to become active and is discarded if the host never opens the note.

  The editor could previously be in a header/footer band, and now also in a note, so the two are one value rather than a flag each: `partEdit` names the single non-body part that is open, and opening a note closes whatever was open before. That removes the state where a band and a note could both claim the caret, and folds the band's first-page variant and originating page into the same value. `PagedEditor` takes `partEdit` in place of `hfEditMode` / `hfEditRId` and reports the open part's selection through `onYrsPartSelectionChange`.

  Everything the body offers behind an open band was already off, and stays off behind an open note for the same reason plus one of its own: the display list indexes images, tables and hyperlinks by band, so a note area has no scope to look one up in and answering `body` would hand back something from behind the note. Selecting a picture, dragging a table edge, the row/column insert affordance and opening a link are therefore inert inside a note; caret movement, word and paragraph selection, drag-selection and Home/End are not. Arrows navigate a note the way they navigate a band — paragraph by paragraph, since the display-list vertical-move query answers body positions only — and the caret never leaves the note it is in.

  Selection geometry comes from the note's own story: `noteCaretRects` is the note twin of `hfCaretRects`, resolving the leading edge of the caret's position and falling back to the trailing edge of the previous one at end of line. A note paints on exactly one page, so unlike a band there is never a second candidate page to choose between. One overlay now paints the caret and highlight for whichever part is open, band or note, and it asks the queries directly — `computeHfCaretRectsFromDisplayList` and `computeHfSelectionRectsFromDisplayList`, which had that overlay as their only caller, are gone from `@betteroffice/docx/layout`.

- 9540e23: Header/footer editing state becomes one value instead of a flag per kind. `partEdit` names the single non-body part the editor has open, folding the band's first-page variant and the page it was opened from into it — three `useState`s in `DocxEditor` become one.

  `PagedEditor` takes `partEdit` in place of `hfEditMode` / `hfEditRId`, and reports the open part's selection through `onYrsPartSelectionChange` in place of `onYrsHfSelectionChange`. `CanvasHfSelectionOverlay` becomes `CanvasPartSelectionOverlay` and queries the display list directly, so `computeHfCaretRectsFromDisplayList` and `computeHfSelectionRectsFromDisplayList` are gone from `@betteroffice/docx/layout`.

  Behaviour is unchanged, with one exception: Escape moves from the header/footer chrome component to the paged area, so it now also closes a band whose chrome never mounted.

- 6be0c18: Bundled metric-compatible fonts ship as `@betteroffice/fonts`, plus `@betteroffice/fonts-cjk` for Chinese, Japanese or Korean, and DOCX uses them only when you hand the module over: `configureDefaultFonts({ fonts })`, or `configureDefaultFonts({ load: () => import('@betteroffice/fonts') })` to keep it in its own chunk. Installing the packages alone does nothing — without that call the engine reaches for no font package, measurement falls back to the browser, and pagination will not match Word. Because `@betteroffice/docx` no longer names `@betteroffice/fonts` anywhere in its published bundle, an esbuild consumer without the optional peer builds again.

### Patch Changes

- e88e5e7: Typing and deletion in the DOCX editor now follow the visible caret after tables, page or column breaks, and block-level content controls.
- 8aa8b37: Edits made in an endnote now reach the saved file. Endnote stories were seeded into the collaborative document but never projected back on save, so every endnote edit was silently replaced by the imported text; footnotes and endnotes now project through one path, and typing in a note or header — including a table cell inside one — marks its root story for the changed-stories-only save.
- 5ff0142: Paragraphs with `lineRule="exact"` or `lineRule="atLeast"` now use Word's measured baseline model
- a2d08a8: The DOCX editor's mouse cursor now reflects what is under it. Hovering typeable text shows the caret cursor and everything else shows the arrow, where the canvas renderer previously painted one arrow over the whole document.

  A position alone could not drive this: the display hit test resolves a caret everywhere on a page that carries any text, margins included, so it cannot say whether a click there would type. `hit_test_regions` now also reports what the point landed on, as `target` on its result — `"text"`, `"image"` or `"none"`. It is the same answer a click acts on and follows the same order: a selectable picture first, since the pointer path picks one before it ever asks for a position, then a run's own box, then the page's typeable area. That area is the authored content box, so the gutter between columns counts like the columns it separates, minus the page's note areas, which this path cannot edit — a click in a footnote lands in the body above it, and the pointer no longer invites one. An area with no positionable text reads as no target for the same reason: a click there jumps the caret to the end of the document.

  A picture is a select target rather than text whatever its wrap mode, because both inline and anchored images carry a document position. One that carries none cannot be selected — a picture watermark — so text painted over it stays readable through it.

  `target` is additive and optional on `DisplayListRegionHit`, so a display-list query answered by an older wasm build still typechecks and reads as "not text". All three query paths — the stateless JSON exports, the session handle, and the resident editing engine — answer from the same resolver, so none can disagree.

  Header and footer bands read as text over their own runs and arrow elsewhere: a double-click there opens the band for editing at that word. While one is open only that band types; a read-only document types nowhere.

- ba687e8: Footnote and endnote text is now reachable by the display hit test. A point that lands in a note area resolves against that note instead of the body line above it, and the pointer shows the caret cursor over a note's glyphs where it used to show the arrow.

  A note is a document of its own, so the region vocabulary had to grow to name one. `hit_test_regions` answers `"footnote"` / `"endnote"` alongside `"body"`, `"header"` and `"footer"`, carrying `noteId` the way a band carries its `rId`: the position then addresses the `fn:{id}` / `en:{id}` story, never the body. A note area stacks several such stories, so the point resolves against the note nearest it rather than the area as a whole — a click can never borrow a position from another note. Nearest counts both axes, because a note area lays its notes out in columns it starts at the same vertical position: telling a pointer in one column from the next takes more than its `y`. An area that paints nothing still owns the click, like an empty header band does, and names no story to route it to.

  That made the display list's own positions load-bearing. A note's primitives kept the note story's range, but anything without one inherited the body range of the reference mark anchoring the note — and under a note region that body number reads as a note-story position, a different document entirely. The reference label leading every note is exactly such a primitive, so this was no edge case: hovering a footnote's number answered with a body position. Those primitives stay unpositioned, and the body anchor rides on the region's `notes` entry, where the accessibility mirror already reads it.

  The label being presentation rather than story content has a visible consequence: it reads as no target, so the cursor is an arrow over a note's number and a caret over its text. A band behaves the same way over its own non-text, and the rule the cursor follows is to under-claim rather than promise typing where a click will not.

  Selection geometry follows the same scoping. `range_rects_in_region` takes the region and the part that owns it as one argument, so a header/footer is named by an `rId` and a note by an id that is never optional — two notes are two unrelated documents whose positions must not mix. `noteRangeRects` is its query-facade twin, and a selection over a note that runs to a bordered paragraph or a table keeps every line it covers rather than stopping at the border.

  The test that pinned note areas as holes in the typeable area is replaced by tests of the new behavior, and the area subtraction it described is gone: a point inside a note is answered by the note before the body is ever asked.

  `region` may now be `"footnote"` or `"endnote"` and `noteId` is optional on `DisplayListRegionHit`, so a display-list query answered by an older wasm build still typechecks. Clicking a note still does nothing — routing a selection into a note story is the editing mode's job, not the layout API's.

- 3533411: The table border toolbar's buttons now do what their names say. `set_cell_borders` replaced a cell's whole `tcPr.borders` object with whatever it was handed, so pressing Top Border on a cell of a bordered table silently deleted that cell's other three rules, and Outside Borders gave every selected cell all four edges — a full grid rather than an outline. It now merges the sides it is given: an omitted side is left alone, `style: "none"` authors an explicit no-border, and a JSON `null` drops the authored side, the same patch convention `set_cell_text_format` already uses.

  Inside Borders wrote `insideH`/`insideV` straight onto each cell. Those keys describe a table's interior edges and have no meaning on a single cell, so nothing rendered them and the grid vanished on screen, while the writer still lifted a complete `w:tblBorders` into the file and the rules reappeared on reopen. `insideH`/`insideV` now resolve per cell to the physical edges interior to the selection, reusing the mapping seeding already applies when it pushes `w:tblBorders` down to cell positions — so an Inside Borders command paints the interior rules of whatever is selected and leaves the outline untouched, and a single-side button applies to the selection's own edge rather than to every cell in it.

  Saving no longer invents table borders no cell carries: the `w:tblBorders` lifted back out of a table is now read from the cells that own each boundary, with `insideH`/`insideV` taken from an interior edge, instead of filling every missing side from the first border found anywhere in the table.

- Updated dependencies [b962e66]
- Updated dependencies [53c583d]
- Updated dependencies [f6af707]
- Updated dependencies [8aa8b37]
- Updated dependencies [5ff0142]
- Updated dependencies [c2e9e69]
- Updated dependencies [a2d08a8]
- Updated dependencies [17f2ead]
- Updated dependencies [7799555]
- Updated dependencies [9540e23]
- Updated dependencies [ba687e8]
- Updated dependencies [43ab7ba]
- Updated dependencies [9540e23]
- Updated dependencies [335bb21]
- Updated dependencies [3533411]
- Updated dependencies [53c583d]
- Updated dependencies [d2e9577]
- Updated dependencies [cd305a5]
- Updated dependencies [6be0c18]
  - @betteroffice/docx@0.1.0
  - @betteroffice/docx-i18n@0.1.0

## 0.0.4

### Patch Changes

- 5c9a482: ArrowUp/ArrowDown move the caret by visual line with persistent goal-X (including across paragraphs, pages, columns, and into tables), and content below tables is clickable and editable again.
- 5c9a482: Collaborative presence: remote collaborator carets and selections render as colored overlays with name flags, anchored by yrs sticky indices so they rebase exactly under concurrent edits; carets follow remote typing instantly by inferring position from document updates.
- 5c9a482: Opening a document now seeds the collaborative session directly in the Rust engine instead of materializing the full TypeScript document model and projecting it; the TS model is built lazily only where the public API still exposes it, and the internal DrawingML host package is dissolved.
- 5c9a482: Remote collaborators' edits no longer move the local viewport: relayouts triggered by remote updates anchor to the topmost visible line via yrs sticky positions and compensate the scroll offset, while caret scrolling fires only for local actions. Anchoring holds across page boundaries too, so text overflowing onto a new page (or pulling back off one) no longer jumps the viewport for either the author or a viewer.
- Updated dependencies [5c9a482]
- Updated dependencies [5c9a482]
- Updated dependencies [5c9a482]
- Updated dependencies [5c9a482]
  - @betteroffice/docx@0.0.4
  - @betteroffice/docx-i18n@0.0.4

## 0.0.3

### Patch Changes

- Updated dependencies [b34bb01]
  - @betteroffice/docx@0.0.3
  - @betteroffice/docx-i18n@0.0.3

## 0.0.2

### Patch Changes

- eed05a6: Fix the published dependency ranges: 0.0.1 shipped the unresolved `workspace:*` protocol for `@betteroffice/docx` and `@betteroffice/docx-i18n`, which made `npm install @betteroffice/docx-react` fail. Ranges are now pinned to concrete versions at publish time.
  - @betteroffice/docx@0.0.2
  - @betteroffice/docx-i18n@0.0.2
