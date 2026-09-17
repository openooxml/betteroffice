---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Measure the `hexagon`, `parallelogram`, `trapezoid` and `octagon` adjust against the shortest side, and pin it at the spec's aspect-scaled maximum, so wide shapes no longer draw their slant or corner at a fraction of the width. The trapezoid defaults to the spec's 25000, and the hexagon honours its `vf` height factor, pinned so the corners stay on the frame.
