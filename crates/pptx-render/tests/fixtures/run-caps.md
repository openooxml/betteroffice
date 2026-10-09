# Run caps

`run-caps.pptx` is a synthetic one-slide deck built on the `run-spacing.pptx` package shell
(same master, layout and theme, slides 2–3 and the media part dropped). It contains no private
corpus material. Every run asks for Arial so a test can register Liberation Sans in its place.

| Shape | `a:rPr/@cap` | Stored text | Drawn text |
| --- | --- | --- | --- |
| 10 | `all` on the run | `Mixed Case Title` | `MIXED CASE TITLE` |
| 11 | `small` on the run | `Small Caps Run` | `SMALL CAPS RUN`, lowercase stretches at 0.8× |
| 12 | `all` on the shape list style's `a:defRPr` | `Inherited Caps` | `INHERITED CAPS` |
| 13 | `none` on the run, `all` on the shape list style | `Kept as authored` | `Kept as authored` |

Casing is a display property, so the fixture pins both projections at once: the story snapshot
still reads the authored text, and saving the deck unmodified re-serializes every part
byte-identically.

Measured on the PPTX corpus tail: `pptarena-051` and `pptarena-052` carry `cap="all"` on a
slide layout's title `a:defRPr`, `pptarena-054` on its master and five slides, and
`pptarena-034` on four slide runs.
