# Chart text run properties

`chart-text-properties.pptx` derives from `chart-space-fill.pptx`, with the
theme's major latin font changed to `Georgia` and its minor left at `Arial`.
Both charts plot the same two revenue series and differ only in their text
properties.

| Scope | Slide 1 (`chart1.xml`) | Slide 2 (`chart2.xml`) |
| --- | --- | --- |
| `c:chartSpace/c:txPr` | `a:latin typeface="+mj-lt"` | none |
| `c:title/c:txPr` | `i="1" spc="300"` | none |
| `c:title/c:tx/c:rich` run | `spc="600"` | none |
| `c:catAx/c:txPr` | `b="1"`, `a:latin typeface="Verdana"` | none |
| `c:legend/c:txPr` | `sz="1400"` | none |

On main `0f474aa5` every run of both charts is drawn in the theme minor font,
upright and untracked: `chart_text_primitive` resolved `+mn-lt` itself and
passed `italic: false` and `letter_spacing_px: 0.0` whatever `c:txPr` declared.
The title's own `spc` never reached the model either, because a title with a
`c:txPr` never looked at its `c:rich`.

With the fix, slide 1 draws the title in italic Georgia at 13px with 8px of
tracking, the category labels in bold Verdana, the legend in 18.67px Georgia
and everything else in Georgia; slide 2 is unchanged, every run Arial at the
sizes the plot geometry defaults to.

Slide 3 keeps `chart-space-fill.pptx`'s third chart and is unused.
