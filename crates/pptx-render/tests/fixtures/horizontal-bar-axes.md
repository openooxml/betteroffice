# Horizontal bar axes

`horizontal-bar-axes.pptx` derives from the public `data-labels.pptx` fixture.
It uses the horizontal bar axis configuration documented in issue #309, with
explicit value bounds to isolate axis layout from automatic scale rounding.
All five slides use categories `Category 1` through `Category 3`, North values
`12, 19, 7`, South values `8, 14, 21`, stacked grouping, and disabled data labels.

| Slide | Variant |
| --- | --- |
| 1 | Horizontal bars, category and value orientations `minMax` |
| 2 | Horizontal bars, category orientation `maxMin` |
| 3 | Column control, both orientations `minMax` |
| 4 | Horizontal bars with a secondary value axis and minor gridlines |
| 5 | Horizontal bars, value orientation `maxMin` |

The chart frame is `(96, 96, 576, 336)` CSS pixels. Rendered with the tracked
Liberation Sans font registered as Arial, slide 1 on main `069e4d66` puts all
five value ticks at x=100, at baselines `127, 195.5, 264, 332.5, 401`. The first
category is at baseline 174.23334, above the third at 356.9. Its purple bar
starts at `(138, 151.4)` with width 126 and height 36.533333.

After the fix, values `0, 10, 20, 30, 40` have x positions
`156, 252.5, 349, 445.5, 542` and baseline 412. Category 1 is at baseline 359.6,
below Category 3 at 188.93333. Category text ends at x=148.91602, before the
plot at x=172. The first purple bar starts at `(172, 338.26666)` with width
115.8 and height 34.133335. North remains `#6254E7`, South `#1FA97A`, text
`#222222`, axes `#666666`, and gridlines `#D9D9D9`.

Reserving an 18px axis header puts the category title at baseline 137, clear of
the chart title at baseline 114. On slide 4, secondary ticks sit at baseline 136
and the plot's right edge is at x=558. Its minor gridlines remain vertical, at
width 0.25.

Slide 2 reverses category order; slide 5 reverses value positions and bars.

Review: [PR #310](https://github.com/openooxml/betteroffice/pull/310).
