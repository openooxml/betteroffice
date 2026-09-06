# Block arrow adjustments

`arrow-adjustments.pptx` replaces slide 1 of the tracked demo deck with block
arrows. Slides 2 and 3 retain the demo's content as controls. `Arrow: Up 35`
uses the exact shape XML from [issue #338](https://github.com/openooxml/betteroffice/issues/338),
including its position, dimensions, rotation, adjustments, and white fill.
The issue names `tpl-result-or-success-slide` but does not attach that deck;
this fixture makes its arrow reproducible without the external corpus.

Only the head adjustment uses the shortest side. The shaft adjustment uses
the full cross axis: width for vertical arrows, height for horizontal arrows.
This follows Microsoft's published [upArrow definition](https://learn.microsoft.com/en-us/answers/questions/2275994/uparrow-is-missing-in-presetshapedefinitions-xml)
and the four arrow definitions in [Apache POI's preset geometry resource](https://github.com/apache/poi/blob/trunk/poi/src/main/resources/org/apache/poi/sl/draw/geom/presetShapeDefinitions.xml).

The issue arrow is 468000 × 1078605 EMU, with `adj1=55713`, `adj2=80407`,
and rotation 1645276/60000 degrees. At 96 DPI its frame is
`(968.1119, 58.812283, 49.133858, 113.239365)` px. Main `8b9f3635` places
the head base at normalized y=0.80407, or 91.0524 px from the tip. The fix
places it at y=0.3488809842, or 39.5071 px. Its shaft edges remain
x=0.221435 and x=0.778565, giving a 27.3739 px shaft. Fill stays `#FFFFFF`.
The exact EMU head length is 376304.76; the old calculation gives 867273.92235.

| Slide 1 shapes | Colour | Expected geometry |
| --- | --- | --- |
| Wide up/down shaft controls, 200 × 50 px | `#FFD966` | 80 px shafts, 25 px heads; unchanged from main |
| Tall left/right shaft controls, 50 × 200 px | `#FFD966` | 80 px shafts, 25 px heads; unchanged from main |
| Long up/down default heads, 50 × 200 px | `#FFFFFF` | 25 px heads, previously 100 px |
| Long left/right heads, 200 × 50 px, `adj2=150000` | `#FFFFFF` | 75 px heads, previously 200 px |
| Square up arrow and rectangle | `#70AD47` | Unchanged controls |

The contributor's shaft scaling shrank the four yellow shafts from 80 px to
20 px. The completed fix preserves them. Unit tests cover that regression,
the issue arrow, defaults and square geometry, every direction, and head
adjustments below zero, above one, and above the shape's length.

The screenshots use the native PPTX raster backend at 96 DPI. Display-list
isolation registers the tracked Liberation Sans, Liberation Serif, Liberation
Mono, Carlito, and Caladea faces in the same order on both revisions, including
their bold and italic variants and common Office aliases. Across all 31
previously tracked decks plus this fixture, 95 of 97 slides are byte-identical.
Only the five white arrow heads here and the purple default right arrow in
`crates/ooxml-drawingml/tests/fixtures/preset-adjustments.pptx`, slide 1, change.
That 100 × 50 px right arrow's head changes from 50 px to 25 px; its four
head-base x coordinates change from 0.5 to 0.75 and its fill stays `#663399`.
Every other display-list field is byte-identical.

No model, writer, or schema change is required. All 32 snapshots serialize
identically and deserialize across revisions. External Python `zipfile`
comparison verifies all 537 package parts on each revision's no-edit save;
ZIP directory entries are not package parts.

| Main | Fixed |
| --- | --- |
| ![Main](arrow-adjustments.before.png) | ![Fixed](arrow-adjustments.after.png) |
