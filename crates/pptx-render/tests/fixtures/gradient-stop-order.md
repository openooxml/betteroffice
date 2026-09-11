# Gradient stop order

[gradient-stop-order.pptx](gradient-stop-order.pptx) is a synthetic, three-slide
deck with seven XML/relationship parts, no external assets, and a 320 × 180 px
slide size. Each slide contains only a gradient background.

| Slide | Stops in XML order | Expected display-list order |
| --- | --- | --- |
| 1, linear | `100000:0000FF`, `0:FF0000` | `0:#FF0000`, `1:#0000FF` |
| 2, radial | `100000:262626`, `0:404040` | `0:#404040`, `1:#262626` |
| 3, linear control | `0:FF0000`, `-0:FF00FF`, `50000:00FF00`, `50000:FFFF00`, `100000:0000FF` | Unchanged, including coincident stops at `0` and `0.5` |

On main `1d0f41d9`, slides 1 and 2 emit positions `[1, 0]`. Native raster output
is uniformly `#0000FF` and `#262626`, respectively. With sorted display-list
stops, slide 1's pixels at `(0,90)`, `(160,90)`, and `(319,90)` are `#EE0011`,
`#7F0080`, and `#1100EE`. Slide 2's center `(160,90)` is `#404040`, its left
edge `(0,90)` is `#292929`, and its corner `(0,0)` is `#262626`.

[Before](gradient-stop-order.before.png) and
[after](gradient-stop-order.after.png) are slide 1's native raster output at 1×.
Slide 3's display list and PNG remain byte-identical. Only
`background.stops` changes on slides 1 and 2; saving without edits preserves
every original ZIP part and its XML stop order.

The renderer regression test covers all four gradient kinds, alpha,
equal-position stops, and source-model preservation. The raster regression
test also feeds an unordered display list directly to the backend and checks
the colors on both sides of the hard edge. Signed zero must compare equal to
positive zero so sorting preserves the XML color order on the control slide.
