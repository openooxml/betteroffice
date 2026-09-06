`text-orientation.pptx` is a synthetic, three-slide repro for #307 and #308, with no private presentation content.

- Slide 1 covers all four flip combinations and two rotated `vert270` bars. Shape 2 uses the issue's exact bar geometry: 639990 × 7718032 EMU at (4601452, -2333065), rotated 90°. With Liberation Sans it changes from 10 lines in a 67.19055 × 810.29205 px box at 90° to one line in an 810.29205 × 67.19055 px box at 0°. Text stays `#17365D`; shape fills stay `#EAF2F8`.
- Slide 2 covers `vert`, `vert270`, layout/master inheritance, an explicit `horz` override, and autofit. Shapes 10 and 11 have left/top/right/bottom insets of 10/20/30/40 px. Their first lines start at (-65, 295) and (135, 275) in the respective text frames.
- Slide 3 is an unchanged control for -15° and 390° horizontal rotations, explicit `horz`, and the existing `eaVert` fallback.

The integration tests exercise the parsed slide/snapshot path, every caret of the horizontalized labels, inheritance, insets, `normAutofit`, and `spAutoFit`. The unit tests also cover master shapes and the separate text hit-test frame.

The insets remain tied to the [shape's bounding rectangle](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.leftinset). Their permutation for vertical text follows the same mapping as [LibreOffice's text-distance handling](https://github.com/LibreOffice/core/blob/master/oox/source/drawingml/textbodyproperties.cxx).

`text-orientation.before.png` and `text-orientation.after.png` are slide 1 rendered on main at `069e4d66` and this branch with the same Liberation fonts. All 62 slides in the 22 pre-existing tracked PPTX fixtures serialize to byte-identical display lists. In this repro only slides 1 and 2 differ, exclusively in text transforms, text boxes, and their layout; slide 3 is byte-identical. All 23 deck snapshots match between revisions and deserialize in both directions; no-edit saves preserve every OPC part byte-for-byte.
