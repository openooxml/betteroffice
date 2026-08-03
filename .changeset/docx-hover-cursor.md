---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

The DOCX editor's mouse cursor now reflects what is under it. Hovering typeable text shows the caret cursor and everything else shows the arrow, where the canvas renderer previously painted one arrow over the whole document.

A position alone could not drive this: the display hit test resolves a caret everywhere on a page that carries any text, margins included, so it cannot say whether a click there would type. `hit_test_regions` now also reports what the point landed on, as `target` on its result — `"text"`, `"image"` or `"none"`. It is the same answer a click acts on and follows the same order: a selectable picture first, since the pointer path picks one before it ever asks for a position, then a run's own box, then the page's typeable area. That area is the authored content box, so the gutter between columns counts like the columns it separates, minus the page's note areas, which this path cannot edit — a click in a footnote lands in the body above it, and the pointer no longer invites one. An area with no positionable text reads as no target for the same reason: a click there jumps the caret to the end of the document.

A picture is a select target rather than text whatever its wrap mode, because both inline and anchored images carry a document position. One that carries none cannot be selected — a picture watermark — so text painted over it stays readable through it.

`target` is additive and optional on `DisplayListRegionHit`, so a display-list query answered by an older wasm build still typechecks and reads as "not text". All three query paths — the stateless JSON exports, the session handle, and the resident editing engine — answer from the same resolver, so none can disagree.

Header and footer bands read as text over their own runs and arrow elsewhere: a double-click there opens the band for editing at that word. While one is open only that band types; a read-only document types nowhere.
