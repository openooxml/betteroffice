---
"@betteroffice/xlsx": minor
"@betteroffice/xlsx-react": minor
"@betteroffice/rust-crates": minor
---

A chart on a worksheet is now a selectable object instead of a picture the click falls through. Every frame publishes each chart's id, rect, clipped hit area and whether it can be repinned, and chrome resolves a press against that frame's own regions, so the answer cannot drift from the pixels. Clicking a chart selects it and outlines it; dragging it, or nudging it with the arrow keys, slides it through the op log as one undoable edit; and the moved anchor — cell and EMU offset alike — is written back into the drawing part on save, synthesising the `colOff`/`rowOff` a drawing omitted rather than saving a move that lost half of itself. Clicking off the chart restores the cell selection, and while a chart is selected the keyboard no longer reaches the cells hidden behind it.

A chart the renderer could not draw is selectable too: it degrades to a placeholder but still occupies its space, so it stays an object that can be picked up and moved out of the way.

Two limits worth stating. A chart pinned by an absolute anchor can be selected but not moved: its position lives in attributes the writer cannot rewrite, and the frame reports this as `movable: false` so the UI never offers the drag. And moving a chart is a standalone-session edit — chart state syncs as one blob per sheet, so a collaborative session refuses it exactly as it refuses freeze panes and hyperlinks.

Minor rather than patch: `DisplayList.charts` changes its element contract. The elements gain `id`, `rect`, `clip` and `movable` beside `placeholder`, and lose `Eq` on the Rust side, so anything that *constructs* one must be updated — `ChartA11yAttrs` survives only as an alias for readers. `Op::SetChartAnchor` is a new variant and breaks exhaustive matches on `Op`.
