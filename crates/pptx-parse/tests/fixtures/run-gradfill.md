`run-gradfill.pptx` is a generated regression deck licensed with this repository.
It provides a public reproduction for PR #299.

The deck has two slides at 960 × 720 CSS pixels. Slide 1 contains these labels
on `#A81020` bands; slide 2 repeats only the three controls.

| Slide 1 shape ID | Label | Main color | Fixed color |
| --- | --- | --- | --- |
| `slide:0:256:shape:0` | RGB gradient | `#505050` | `#FFFFFF` |
| `slide:0:256:shape:1` | Theme gradient (`lt1`) | `#505050` | `#FFFFFF` |
| `slide:0:256:shape:2` | Modified gradient (`204060`, 50% tint) | `#505050` | `#90A0B0` |
| `slide:0:256:shape:3` | Unsorted green/red ramp | `#505050` | `#00FF00` |
| `slide:0:256:shape:4` | Adjacent white/red and white/blue ramps | `#505050` | `#FFFFFF` |
| `slide:0:256:shape:5` | Solid control | `#FFFFFF` | `#FFFFFF` |
| `slide:0:256:shape:6` | Inherited control | `#505050` | `#505050` |
| `slide:0:256:shape:7` | Empty gradient | `#505050` | `#505050` |

The six authored gradient runs become five display-list text runs because the
adjacent runs have the same modeled style. Stops are deliberately stored in
reverse position order. The lowest valid stop is a fallback for multicolor
ramps; a single-color gradient renders exactly. The writer retains the authored
ramps when text or other formatting changes, including adjacent runs separated
by a soft line break, and replaces the fill when the resolved color changes.

Review: [PR #299](https://github.com/openooxml/betteroffice/pull/299).

Regression tests cover exact stop selection, invalid and empty stop lists, solid
fill priority, rendering, and saves after bold formatting, ASCII/Unicode text
insertion, and recoloring. Multicolor ramps remain in the saved XML after
non-color edits.
