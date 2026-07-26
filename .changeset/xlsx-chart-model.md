---
"@betteroffice/xlsx": patch
"@betteroffice/rust-crates": patch
---

Charts are modeled. A worksheet's charts are read through the drawing that anchors them — sheet relationships to `xl/drawings/drawingN.xml` to `xl/charts/chartN.xml` — and every `c:f` reference and `xdr:` anchor becomes part of the workbook model. Row and column edits, sheet renames and sheet removals now rewrite those references the same way they rewrite cell formulas, hyperlink locations and defined names: references shift, ranges clip, wholly deleted ranges collapse to `#REF!`, and a rename rewrites the sheet qualifier. Anchors follow the grid according to `editAs` — `twoCell` moves and resizes, `oneCell` moves without resizing, `absolute` does neither. Saving patches the moved `c:f` values and anchor indices in place, so styling, 3-D views, data labels, trendlines and cached values survive byte for byte.

Because those references are now rewritten, sheet rename, sheet removal and row or column edits are no longer refused on a workbook carrying a chart. Pivot tables still refuse them: pivot caches are not modeled.

The collaboration authority moves to schema version 6 so chart state syncs between peers. Snapshots written at versions 3 through 5 upgrade when they are read.

Known limits. A chart reference the reference rewriter cannot express — anything beyond a cell, a range, a whole row or column, or a parenthesized union of those — refuses the edit rather than being left on the pre-edit addresses, the same way an unrewritable defined name does. A `oneCell` anchor keeps its cell span rather than its exact EMU size when a row or column is inserted inside that span. Chart parts are read for their references and anchors; rendering them is not wired up yet.
