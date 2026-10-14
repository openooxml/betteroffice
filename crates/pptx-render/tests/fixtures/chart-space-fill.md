# Chart-space fill and axis lines

`chart-space-fill.pptx` derives from `data-labels.pptx`. Each slide carries one
revenue chart on a dark `#1F2933` slide background, with the North series
painted `bg1` (white) and South `accent2` (`#1FA97A`). The three chart parts
vary only their `c:chartSpace/c:spPr` and the `c:spPr/a:ln` of each axis.

| Slide | Chart space | Category axis line | Value axis line |
| --- | --- | --- | --- |
| 1 | `a:pattFill prst="smGrid"` `01C4BF` over `01BABC` | `w="9525"`, `tx1` at `lumMod 15%`/`lumOff 85%` | `a:noFill` |
| 2 | `a:noFill` | none declared | none declared |
| 3 | `a:solidFill` `lt2` (`#F5F7FA`) | `w="19050"`, `BFBFBF` | `w="6350"`, `accent1` |

The chart frame is `(96, 96, 576, 336)` CSS pixels and the plot spans
`x = 138..558`, `y = 124..398`. On main `9274a2ba` every slide opens with a
`#FFFFFF` rectangle over the frame and strokes both plot edges `#666666` at 1px,
so the white North bars vanish on slides 1 and 3 and the dark slide never shows
through on slide 2.

With the fix:

- Slide 1 grounds the frame in `#01BFBD`, the mean of the two pattern colours,
  draws no value-axis rule, and strokes the bottom rule `#D9D9D9` at 1px.
- Slide 2 draws no ground rectangle at all; the two `#666666` rules stay.
- Slide 3 grounds the frame in `#F5F7FA`, strokes the left rule `#6254E7` at the
  1px hairline floor (`6350` EMU is 0.67px) and the bottom rule `#BFBFBF` at 2px.

The title, six value ticks, three categories, two axis titles and two legend
entries account for 14 text primitives on every slide, and the six bars keep the
same geometry on all three.
