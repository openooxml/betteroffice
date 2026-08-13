---
"@betteroffice/xlsx": minor
"@betteroffice/rust-crates": minor
---

Dragging a chart works in a collaborative session. It used to fail outright with "structural operations are unavailable in collaborative mode", because repinning a chart sat in the structural list beside inserting a row — so the feature could not work at all in the mode the collaborative demo runs in.

A chart anchor is not workbook structure. What is frozen is now what a chart *is*: the drawing anchor it sits at, the part behind it, the references it reads, and the shape of that anchor — its kind, its `editAs` mode, a one-cell extent, an absolute position. Where a grid-anchored chart sits travels through the shared document like any other edit, and a peer picks it up the same way it picks up a cell.

Two things follow from letting an anchor merge, and both are worth knowing.

An arriving anchor is judged only on what is true of it whatever grid it sits over: offsets in range, corners in order, extents positive. Whether it resolves to a rectangle with room to draw depends on the column widths it spans, and those are replicated too — so that question is settled where the chart is drawn, not where the update arrives. Collapsing the columns under a chart while someone else drags it leaves both replicas holding the same workbook; the chart simply has nowhere to render, which is what collapsing those columns asked for.

One drawing anchor is one element however many sheets point at it, so a drag repins every sheet holding it, and a workbook read back from the shared document always shows them agreeing. Where two sheets have been written separately, the first in sheet order supplies the anchor.

Chart state is one value per sheet, so two people dragging different charts on the same sheet keep only one of the two drags. The replicas agree on which; nothing is corrupted and no one is disconnected. Dragging the same chart at once already resolves to a single anchor.
