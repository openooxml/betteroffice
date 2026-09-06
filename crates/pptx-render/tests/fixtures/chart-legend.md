# Chart legends and titles

`chart-legend.pptx` derives from the tracked
`crates/pptx-parse/tests/fixtures/chart-deck.pptx`. It retains the first chart,
its two series, theme and axes, and removes the deliberately broken chart frame.
It contains no external presentation content.

The chart frame is `(96, 96, 576, 336)` CSS pixels. Narrow cases use width 200.
North has values `[12, 19, 7]` and colour `#6254E7`; South has values
`[8, 14, 21]` and colour `#1FA97A`. The title is `Revenue`.

| Slide | Case |
| --- | --- |
| 1 | Bottom legend |
| 2 | Top legend |
| 3 | Right legend, title alignment control |
| 4 | Left legend, title alignment control |
| 5 | Bottom legend, narrow frame, long series names |
| 6 | Top legend with a 30pt legend font |
| 7 | Right legend with all chart and axis titles removed; unchanged control |
| 8 | Bottom legend on a horizontal bar chart |
| 9 | Bottom legend, narrow frame, ten `W` and ten `M` characters |
| 10 | Bottom legend, narrow frame, multi-line series names |

`chart_legend.rs` checks placement, separation, complete wrapped text, shaped
widths and title alignment. The shared geometry tests additionally check pie
and radar regions, column legends and returned plot width.

The before image uses main at `069e4d66`; the after image uses this change.
Both render slide 1 with the same Liberation Sans regular and bold fonts under
the deck's Calibri family, at 96 dpi. The title's shaped centre moves from
`131.095` to the frame centre, `384`. The legend swatches move from the right
column at `(574, 132)` and `(574, 147)` to the bottom band at `y=417`.
The plot grows from 420 to 516 pixels wide and shrinks from 274 to 252 pixels
high. The retained 8px outer margin accounts for the 96px width increase.

The long-label images show slide 5 before and after the review corrections.
The contributor's single row overlaps the two labels and clips the second at
the chart edge; the corrected layout gives each entry a complete row.
