`paragraph-spacing.pptx` is a synthetic, one-slide repro for paragraph spacing
built from the public demo package. It contains no third-party presentation
content.

Every body has zero insets and three unwrapped paragraphs, so the gap between
paragraphs is the only thing that varies. The master `bodyStyle` sets
`spcBef` to 10 pt; the layout's second body placeholder adds `spcAft` 18 pt.

| Shape | Source | Font size | Expected gap |
| --- | --- | --- | --- |
| 2 | Master `spcBef` 10 pt | 18 pt | 13.33 px |
| 3 | Master 10 pt before plus layout 18 pt after | 18 pt | 37.33 px |
| 4 | Direct `spcBef` 50% | 32 pt | 21.33 px |
| 5 | Direct `spcBef` 0% over the master's 10 pt | 18 pt | 0 px |
| 6 | Direct `spcBef` 6 pt and `spcAft` 6 pt | 18 pt | 16 px |
| 7 | Direct `spcAft` 12 pt, bottom anchored | 18 pt | 16 px |
| 8 | No spacing | 18 pt | 0 px |

No space is opened above the first paragraph or below the last: shapes 2 to 5
start at the top of their box and shape 7 ends at the bottom of its box.
LibreOffice renders all seven at those pitches.

Percentages measure the paragraph's text size, per
[SpacingPercent](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spacingpercent);
point units are hundredths of a point, per
[SpacingPoints](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spacingpoints).
