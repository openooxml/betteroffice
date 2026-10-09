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
