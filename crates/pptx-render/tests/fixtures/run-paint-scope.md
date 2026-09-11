# Adjacent run paint

`run-paint-scope.pptx` is a synthetic three-slide deck with explicit Arial text.
Render with the tracked Liberation Sans regular face registered as Arial to
exercise style changes that share one fallback font ID.

| Slide | Content | Expected result |
| --- | --- | --- |
| 1 | Gold/black/gold, black/gold/black, and separate bold, italic, underline, and size runs | Each run retains its paint; `large` is 48 px and the other text is 32 px |
| 2 | Alternating colours in wrapped and justified paragraphs | Colour boundaries survive within each line; glyph and caret positions stay unchanged |
| 3 | Identically styled adjacent text | Display list stays byte-identical |

Gold is `#A99A72`; black is `#000000`. Every slide also inherits a master
textbox containing two identically styled runs, `same ` and `style`, which
must remain one positioned run. This exercises inherited source runs before
the editor can coalesce identical text attributes.

On main, slide 1 paints `gold black gold` entirely gold and `black gold black`
entirely black. Its last row drops the bold and italic flags, underline, and the 48 px size.
The adjacent before/after PNGs render this slide with the same registered font.

The tracked demo deck also reproduces the bug on slide 3: the three numbered
step labels should retain their accent colours and 16 px size, while the
following descriptions use `#101828` at 20 px.

Run the focused regression tests with
`cargo test -p betteroffice-pptx-render --lib adjacent_runs`.
