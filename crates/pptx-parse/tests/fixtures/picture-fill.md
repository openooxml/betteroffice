`picture-fill.pptx` is a generated, text-free regression deck licensed with this repository.

Its single slide is 480 × 270 CSS pixels and holds three shapes, all filled from the same
64 × 64 bitmap at `ppt/media/image1.png` whose pixels are RGB `(4x, 4y, 40)`.

| Shape | Box (px) | Fill | Before | After |
| --- | --- | --- | --- | --- |
| Freeform ring | `(20, 20, 200, 200)` | `a:blipFill` + `a:stretch` | Nothing drawn | Bitmap clipped to the outer square, counter punched out |
| Banded ellipse | `(250, 20, 200, 200)` | `a:blipFill`, `srcRect l="10000" r="20000"`, `fillRect l="-50000"` | Blue outline only | Bitmap clipped to the ellipse, source cropped to `(1/3, 0.8)` horizontally |
| Tiled rectangle | `(20, 230, 100, 30)` | `a:blipFill` + `a:tile` | Nothing drawn | Unchanged: tiling is not implemented |

The ring is one `a:path` in a `200 × 200` space holding two contours: an outer square wound
clockwise and an inner square, from `(50,50)` to `(150,150)`, wound counter-clockwise. The
nonzero winding rule of both backends is what leaves the counter unpainted.

The banded ellipse pins the fill-rectangle arithmetic. `fillRect l="-50000"` stretches the
image into a band 1.5 times the box width starting half a box width to its left, so the box
shows the source from `0.5 / 1.5` of the way in; composed with the `srcRect` that keeps
`0.1 … 0.8`, the visible source starts at `0.1 + 0.7 / 3` and ends at `0.8`.
