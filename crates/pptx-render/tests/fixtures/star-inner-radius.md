# Star presets

`star-inner-radius.pptx` replaces slide 1 of the tracked demo deck with sixteen
star presets on a white ground. Slides 2 and 3 are unchanged controls. Every
shape is named `<preset> <adj> <width>x150`.

| Row | Colour | Shapes, 150 × 150 px unless noted |
| --- | --- | --- |
| 1 | `#F4B400` | `star5` with no `adj`, then `adj` 19098, 30000, 50000, 80000, 10000 |
| 2 | `#4285F4` | `star4`, `star6`, `star7`, `star8`, `star10`, `star12`, no `adj` |
| 3 | `#0F9D58` | `star16`, `star24`, `star32`, no `adj`; `star5` at 350 × 150, no `adj` |

The [ECMA-376 preset definitions](https://github.com/apache/poi/blob/trunk/poi/src/main/resources/org/apache/poi/sl/draw/geom/presetShapeDefinitions.xml)
give every star `a = pin 0 adj 50000` and an inner radius of `iwd2 = swd2 * a / 50000`,
so the inner-to-outer ratio is `2 * adj / 100000`. Each preset has its own default:
12500 for `star4`, 19098 for `star5`, 28868 for `star6`, 34601 for `star7`,
42533 for `star10`, and 37500 for the rest. `star5`, `star6`, `star7` and `star10`
also scale their outer radii by `hf`/`vf` so their points reach the frame.

At 19098 the ratio is 0.38196, a regular pentagram's: each inner vertex lies on the
line joining two outer points. `adj` 50000 draws a decagon, and 80000 pins to the same.
The first shape's bottom points land at y = 210 px, the frame's bottom edge.

Review: [PR #466](https://github.com/openooxml/betteroffice/pull/466).

The comparison uses LibreOffice 26.8 and native PPTX raster output at 96 DPI.
