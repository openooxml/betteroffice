# Gradient outlines

Repro: [gradient-outline.pptx](../../../pptx-parse/tests/fixtures/gradient-outline.pptx), slide 1.

Main `069e4d66bf749869ad581114dd3e4e4c721ed07f` emits only the solid control's stroke.
The rectangle at `(96, 96, 768, 192)` and horizontal line at `(96, 336, 768, 0)`
have no stroke. Both now carry an 8 px linear gradient at 0 degrees with stops
`0: #C00000`, `0.5: #FFC000`, and `1: #1F7A3D`. Their flat fallback is `#C00000`.
The control at `(96, 432, 768, 192)` remains an 8 px `#C00000` stroke with no paint field.

The images use the same Liberation Sans font bytes and the Rust raster backend:

| Before | After |
| --- | --- |
| ![Before](gradient-outline.before.png) | ![After](gradient-outline.after.png) |

At `(96, 336)`, `(480, 336)`, and `(863, 336)`, main paints white. The fixed
renderer paints `(192, 0, 0)`, `(255, 192, 0)`, and `(31, 122, 61)`, respectively.
Exactly 21,504 pixels change: 15,360 on the rectangle outline and 6,144 on the line.
The control and every pixel outside those two strokes remain identical.
