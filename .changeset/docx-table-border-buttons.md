---
"@betteroffice/rust-crates": patch
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

The table border toolbar's buttons now do what their names say. `set_cell_borders` replaced a cell's whole `tcPr.borders` object with whatever it was handed, so pressing Top Border on a cell of a bordered table silently deleted that cell's other three rules, and Outside Borders gave every selected cell all four edges — a full grid rather than an outline. It now merges the sides it is given: an omitted side is left alone, `style: "none"` authors an explicit no-border, and a JSON `null` drops the authored side, the same patch convention `set_cell_text_format` already uses.

Inside Borders wrote `insideH`/`insideV` straight onto each cell. Those keys describe a table's interior edges and have no meaning on a single cell, so nothing rendered them and the grid vanished on screen, while the writer still lifted a complete `w:tblBorders` into the file and the rules reappeared on reopen. `insideH`/`insideV` now resolve per cell to the physical edges interior to the selection, reusing the mapping seeding already applies when it pushes `w:tblBorders` down to cell positions — so an Inside Borders command paints the interior rules of whatever is selected and leaves the outline untouched, and a single-side button applies to the selection's own edge rather than to every cell in it.

Saving no longer invents table borders no cell carries: the `w:tblBorders` lifted back out of a table is now read from the cells that own each boundary, with `insideH`/`insideV` taken from an interior edge, instead of filling every missing side from the first border found anywhere in the table.
