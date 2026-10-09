`custom-geometry.pptx` is a generated, text-free regression deck licensed with this repository.

Slide 1 contains two targeted shapes and three controls. Slide 2 contains only presets.

| Shape | Before | After |
| --- | --- | --- |
| Mixed paths, box `(20,20,200,100)` px | Red rectangle with blue outline | Four independent paths, with line, cubic, quadratic and close commands |
| Quarter ellipse, box `(260,20,200,100)` px | Green rectangle with blue outline | Elliptical arc from `(440,40)` to `(360,70)` px, then line and close |
| Missing path list / unresolved guide | Amber rectangles | Identical fallback rectangles |
| Preset ellipse | Purple ellipse | Identical preset path |

The mixed shape's first path uses a `200 × 100` coordinate space against a `1905000 × 952500` EMU transform. Its first points normalize to `(0.1,0.1)` and `(0.9,0.1)`, which render at `(40,30)` and `(200,30)` px. Its remaining paths use `400 × 200`, `100 × 50`, and `200 × 100` spaces. Their paints are respectively stroke only, fill only, and neither. The fill is `#DC2626`; the stroke is `#2563EB`, 2 px wide.

The quarter ellipse uses radii `(80,30)` in a `200 × 100` path space and a 90-degree sweep. Its normalized cubic controls are `(0.9,0.3656854249492381)` and `(0.7209138999323174,0.5)`, with endpoint `(0.5,0.5)`.

Review: [PR #285](https://github.com/openooxml/betteroffice/pull/285).

The version-2 update fixture was regenerated with current `origin/main` at
`54fdaa00c8242d58db61418ac3bc3b2ad6d50cb4`, using client ID 285. It is stored at
`crates/pptx-edit/tests/fixtures/deck-custom-schema-v2.update.bin`. The fixture
generator uses main’s legacy parser and defaults before stamping v2. It contains
the parsed model without custom paths and exercises migration to the current
schema. Restoring an old update without source bytes retains its historical
fallback geometry; attaching the original deck reparses the custom paths.
See the [generator and compatibility details](../../../pptx-edit/tests/fixtures/README.md).

The parser uses the same normalized command representation and 2,048-command bound as the DOCX custom-geometry parser. PPTX preserves individual path painting and rejects an unsupported path as a whole. It converts DrawingML polar angles to ellipse parameters before producing cubic curves; see [Apache POI's DrawingML angle convention](https://github.com/apache/poi/blob/trunk/poi/src/main/java/org/apache/poi/sl/draw/geom/ArcToCommand.java). Guide formulas still use the rectangle fallback.
