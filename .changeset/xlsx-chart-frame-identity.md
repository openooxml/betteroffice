---
"@betteroffice/xlsx": minor
"@betteroffice/rust-crates": minor
---

A chart frame is now addressed by the drawing anchor it sits at rather than by the chart part behind it. Two anchors in one drawing may point at one `chart1.xml`, and both frames used to publish the same `ChartRegion.id`: hit-testing the second one and dragging it moved the first, the selection outline snapped to the first, and the workbook then refused to save at all because one part matched two charts. Each frame now carries its own id, so a drag moves the frame that was picked up and a save writes each anchor back where it sits.

`ChartRegion.id` is `"<drawing part>#<anchor index>"`, opaque, and never a chart part path. `WorkbookHandle.moveChart`, `Workbook::move_chart` and `Op::SetChartAnchor` all take it.

An anchor index names a position in a drawing, not an object, so an op stored against one outlives the structure it was recorded against: another editor that adds, reorders or removes an anchor renumbers every ordinal after it. `Op::SetChartAnchor` is therefore a compare-and-set — it carries `part` and `from`, the chart part and anchor the frame held when the op was recorded, beside `to` — and replay onto a drawing whose anchors have shifted is refused rather than landing on whichever frame now sits at that ordinal. The refusal is the typed `Error::ChartFrameShifted`, so a host replaying a stored log can drop that op and carry on instead of matching on prose. Two frames pinned to byte-identical anchors *and* backed by one chart part are the only case left that the guard cannot tell apart, and they are interchangeable: same part, same rect.

A worksheet that reaches one drawing through two relationships now walks it once; following each separately emitted every anchor twice, and the twins then shared an ordinal.

The refusal that protects a shared chart part is unchanged in strength, only re-expressed per anchor: a save still demands that every frame a sheet carries match exactly one frame the source sheet was read with, that none be dropped, that none be carried twice, and that sheets sharing a part or a drawing agree on what it holds.
