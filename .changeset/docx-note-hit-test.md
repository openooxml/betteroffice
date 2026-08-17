---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Footnote and endnote text is now reachable by the display hit test. A point that lands in a note area resolves against that note instead of the body line above it, and the pointer shows the caret cursor over a note's glyphs where it used to show the arrow.

A note is a document of its own, so the region vocabulary had to grow to name one. `hit_test_regions` answers `"footnote"` / `"endnote"` alongside `"body"`, `"header"` and `"footer"`, carrying `noteId` the way a band carries its `rId`: the position then addresses the `fn:{id}` / `en:{id}` story, never the body. A note area stacks several such stories, so the point resolves against the note nearest it rather than the area as a whole — a click can never borrow a position from another note. Nearest counts both axes, because a note area lays its notes out in columns it starts at the same vertical position: telling a pointer in one column from the next takes more than its `y`. An area that paints nothing still owns the click, like an empty header band does, and names no story to route it to.

That made the display list's own positions load-bearing. A note's primitives kept the note story's range, but anything without one inherited the body range of the reference mark anchoring the note — and under a note region that body number reads as a note-story position, a different document entirely. The reference label leading every note is exactly such a primitive, so this was no edge case: hovering a footnote's number answered with a body position. Those primitives stay unpositioned, and the body anchor rides on the region's `notes` entry, where the accessibility mirror already reads it.

The label being presentation rather than story content has a visible consequence: it reads as no target, so the cursor is an arrow over a note's number and a caret over its text. A band behaves the same way over its own non-text, and the rule the cursor follows is to under-claim rather than promise typing where a click will not.

Selection geometry follows the same scoping. `range_rects_in_region` takes the region and the part that owns it as one argument, so a header/footer is named by an `rId` and a note by an id that is never optional — two notes are two unrelated documents whose positions must not mix. `noteRangeRects` is its query-facade twin, and a selection over a note that runs to a bordered paragraph or a table keeps every line it covers rather than stopping at the border.

The test that pinned note areas as holes in the typeable area is replaced by tests of the new behavior, and the area subtraction it described is gone: a point inside a note is answered by the note before the body is ever asked.

`region` may now be `"footnote"` or `"endnote"` and `noteId` is optional on `DisplayListRegionHit`, so a display-list query answered by an older wasm build still typechecks. Clicking a note still does nothing — routing a selection into a note story is the editing mode's job, not the layout API's.
