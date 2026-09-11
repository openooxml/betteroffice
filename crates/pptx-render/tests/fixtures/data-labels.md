# Chart data labels

`data-labels.pptx` derives from `pptx-parse/tests/fixtures/chart-deck.pptx`.
Each slide contains its revenue chart, with the broken chart frame removed.
The five chart parts vary only their `c:dLbls` declarations.

| Slide | Declaration | Expected data labels |
| --- | --- | --- |
| 1 | All six group `show*` switches are `0` | None |
| 2 | All six switches are `0` on each series; no group declaration | None |
| 3 | All six group switches are `0`; each series enables `showCatName` | `Q1`, `Q2`, `Q3` for each series |
| 4 | All six group switches are `0`; North point 1 enables `showVal` | `19` |
| 5 | Group enables `showVal`; each series sets all six switches to `0` | None |

The chart frame is `(96, 96, 576, 336)` CSS pixels. North's three bars and
legend swatch are purple (`#6254E7`); South's are green (`#1FA97A`). The title,
five value-axis ticks, three categories, two axis titles, and two legend entries
account for 13 text primitives. These and every non-text primitive remain intact.

On main `1d0f41d9`, slides 1, 2, and 5 contain 19 chart text primitives instead
of 13. Slide 4 contains 19 instead of 14. Slide 3 remains at 19. The unwanted
text is `#222222`; removing it does not change any bar's path, bounds, or paint.

`data-labels-before.png` and `data-labels-after.png` render slide 1 on main and
the fixed branch using the tracked Liberation Sans font as Arial, at 1× scale.
