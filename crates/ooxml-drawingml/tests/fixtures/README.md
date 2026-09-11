`preset-adjustments.pptx` and `preset-adjustments.docx` replace the first slide or body of the tracked demo documents with synthetic preset shapes. PPTX slides 2 and 3 remain unchanged controls. The bottom row of slide 1 contains ten other preset controls with ordinary adjustment values.

At 96 DPI, the point depth must be `min(max(adj / 100000, 0) * min(width, height), width)`. A missing adjustment uses `50000`. This follows the chevron and homePlate definitions in the [ECMA-376 geometry addendum](https://ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip): `maxAdj = 100000 * w / ss`, `a = pin 0 adj maxAdj`.

| PPTX shape | Colour | Size in pixels | Raw adjustment | Expected point depth |
| --- | --- | --- | --- | --- |
| Wide Chevron Default | `FF0000` | 171.3 × 55.6 | default | 27.8 px |
| Wide HomePlate Adj25 | `FF00FF` | 280.9 × 37.5 | 25000 | 9.375 px |
| Square HomePlate Adj75 | `663399` | 100 × 100 | 75000 | 75 px |
| Square Chevron Adj75 | `CC0066` | 100 × 100 | 75000 | 75 px |
| Wide Chevron Adj200 | `008080` | 200 × 50 | 200000 | 100 px |
| Wide HomePlate Adj200 | `800000` | 200 × 50 | 200000 | 100 px |
| Wide Chevron Adj600 | `808000` | 200 × 50 | 600000 | 200 px |
| Tall HomePlate Adj200 | `000080` | 50 × 200 | 200000 | 50 px |

The DOCX exercises the same parser boundary with explicit raw guide values, including `50000`, `75000`, `200000`, and `600000`. It retains an explicitly adjusted square chevron and a rectangle as controls.

The square `75000` cases must have a normalized first line endpoint at `(0.25, 0)`. The wide `200000` cases must have that endpoint at `(0.5, 0)`. A fixed `0.5` adjustment cap or a global `1.0` normalized cap breaks these cases.

The screenshots render slide 1 through the native PPTX raster backend with the tracked Liberation Sans font. Before uses main commit `387f2392c44e31e459264663fafd65581c8346a6`; after uses the corrected geometry.

| Before | After |
| --- | --- |
| ![Before](preset-adjustments.before.png) | ![After](preset-adjustments.after.png) |
