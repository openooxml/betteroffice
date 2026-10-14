# Symbol-font bullets

`symbol-bullets.pptx` is a synthetic one-slide deck on the same package shell as `run-caps.pptx`;
it contains no private corpus material. Each shape sets one `a:buFont`/`a:buChar` pair and asks for
Arial as its text face, so a test can register Liberation Sans in its place.

| Shape | `a:buFont` | `a:buChar` | Slot | Drawn |
| --- | --- | --- | --- | --- |
| 10 | Wingdings | `§` U+00A7 | `square4` | `▪` |
| 11 | Wingdings 3 | U+F075 | `trianglert` | `►` |
| 12 | Wingdings | `l` U+006C | `circle6` | `●` |
| 13 | Wingdings | U+F0FC | `checkbld`, no covered equivalent | `•` |
| 14 | Arial | `»` U+00BB | not a symbol face | `»` |

The slots and their shapes are the faces' own glyph names; the size within a shape class is the
glyph's ink width against the candidates' ink widths in Liberation Sans at the same em. Shape 11's
choice is confirmed against PowerPoint's render of `pptarena-018` page 3, where the same
Wingdings 3 slot draws a solid right-pointing triangle, and shape 10's against `pptarena-054`
page 3, where Wingdings `square2` draws a small filled square.
