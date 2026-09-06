`line-spacing.pptx` is a synthetic, one-slide repro for PR #314 built from the public demo package. It contains no third-party presentation content.

All text is `#2040B0` in Arial (Liberation Sans for the regression checks). Each text box has two lines and a top anchor.

| Shape | Source | Font size | Expected pitch |
| --- | --- | --- | --- |
| 2 | Master title style, 80%, `compatLnSpc=1` | 32 pt | 40.96 px |
| 3 | Master title style, 80%, `compatLnSpc=0` | 32 pt | 39.25 px |
| 4 | Direct paragraph, 72 pt exact | 32 pt | 96 px |
| 5 | Layout paragraph, 120%, inherited `compatLnSpc=1` | 32 pt | 61.44 px |
| 6 | Direct paragraph overrides layout, 150% | 32 pt | 76.8 px |
| 7 | No explicit or inherited spacing | 24 pt | 36.796875 px |
| 8 | Shape list style, 60%, `compatLnSpc=1` | 18 pt | 17.28 px |

Expanded spacing leaves the first baseline at the font's natural ascent; compressed spacing scales the line box. The integration tests also check exact spacing at 50% autofit and zero spacing overriding the master.

`line-spacing.before.png` is rendered on main `2c90c17f`; `line-spacing.after.png` uses this branch. The comparison checks every display-list field: only line y, height, baseline and glyph y-offsets change; the control shape is identical.

The companion `crates/pptx-edit/tests/fixtures/deck-schema-v9-line-spacing.update.bin` was emitted by main `2c90c17f` after inserting `Edited ` into the first story. It proves schema-10 migration preserves edits and source attachment restores omitted spacing permanently, including after reopening without the source.

The valid spacing ranges follow the [Open XML SDK DrawingML schema](https://github.com/dotnet/Open-XML-SDK/blob/main/data/schemas/schemas_openxmlformats_org_drawingml_2006_main.json). Point units are documented in [SpacingPoints](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spacingpoints).
