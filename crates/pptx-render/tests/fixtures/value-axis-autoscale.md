# Value axis autoscale

`value-axis-autoscale.pptx` derives from the public `data-labels.pptx` fixture.
Only `ppt/charts/chart1.xml` (slide 1) changes: the bar chart is `stacked` with
`overlap` 100, North is `2.9, 3, 4.1`, South is `5.8, 5.9, 8.2` (totals
`8.7, 8.9, 12.3`), and the value axis keeps only
`<c:scaling><c:orientation val="minMax"/></c:scaling>` — no `c:min`, no `c:max`,
no `c:majorUnit`, no `c:numFmt`. Slides 2 to 5 are unchanged copies.

The chart frame is `(96, 96, 576, 336)` CSS pixels. Rendered with the tracked
Liberation Sans font registered as Arial, slide 1 on main `3d95068f` labels the
value axis `0, 3.1, 6.1, 9.2, 12.3` at baselines `401, 332.5, 264, 195.5, 127`,
and the tallest stack (Q3 South, `#1FA97A`) starts at y=124, on the plot's top
edge.

After the fix the labels are `0, 2, 4, 6, 8, 10, 12, 14` at baselines `401` down
to `127` in steps of `39.142857`, and the Q3 South bar starts at y=157.27142,
between the tick marks for 12 (y=163.14285) and 14 (y=124).
