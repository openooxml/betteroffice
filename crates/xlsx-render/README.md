# betteroffice-xlsx-render

Grid geometry and the target-agnostic display list: turns a workbook viewport
into draw commands. Never imports canvas, DOM, or any raster backend.

```rust
use xlsx_render::{Viewport, build_display_list};

let viewport = Viewport { x: 0.0, y: 0.0, width: 1200.0, height: 800.0 };
let display_list = build_display_list(&workbook, sheet_id, &viewport);
```

`GridGeometry` resolves column widths, row heights, and frozen panes into
pixels. `build_display_list` walks the visible range and emits `DrawCmd`s with
resolved fills, borders, alignment, and formatted text. `display_text` applies
the number format for a single cell.

`build_display_list_with_ghosts` overlays `GhostEdit`s, which is how pending
agent proposals paint as in-cell tracked-change previews.

`region` computes the viewport for a range or for the sheet's used range —
useful for exporting a fixed region rather than a scroll position.

The browser replays the display list onto Canvas2D;
[betteroffice-xlsx-raster](https://crates.io/crates/betteroffice-xlsx-raster)
paints the same list to PNG server-side.

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
