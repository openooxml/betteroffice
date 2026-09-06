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

The contributor's original geometry placed the category title at baseline 119,
overlapping the chart title at baseline 114. Reserving an 18px axis header puts
it at baseline 137. On slide 4, secondary ticks move from baseline 118 to 136,
and the obsolete 38px right margin is removed: the plot's right edge moves
from x=520 to x=558. Its minor gridlines remain vertical, at width 0.25.

Slide 2 reverses category order; slide 5 reverses value positions and bars.
Slide 3 is byte-identical to main. No model, snapshot format, or writer changes
are needed. External ZIP-part comparison of the no-edit save preserves all
parts, including the chart XML.

`horizontal-bar-axes.before.png` and `horizontal-bar-axes.after.png` show slide 1
on main and the completed branch. The `secondary.before.png` image shows the
contributor's layout of slide 4; `secondary.after.png` shows its completed layout.
