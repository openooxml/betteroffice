# betteroffice-pptx-render

Slide layout and the PPTX display-list compiler. Takes a composed slide —
shapes resolved against their layout and master — and emits the draw commands a
host replays.

The display list carries vector geometry, images, shaped text, caret data, and
hit-test metadata. It is target-agnostic: no canvas, no DOM, no raster backend.
The browser replays it onto Canvas2D, and a native host can paint it anywhere.

`compile_json` is the entry point.

Used by [betteroffice-pptx](https://crates.io/crates/betteroffice-pptx).

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
