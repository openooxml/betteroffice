# Line ends

`line-ends.pptx` replaces slide 1 of the repository's demo deck with synthetic
line shapes. Slides 2 and 3 are unchanged. No external template is required.

The first five rows exercise `triangle`, `arrow`, `stealth`, `diamond`, and
`oval`. Each row uses width/length pairs `sm/lg`, `med/sm`, and `lg/med` at the
head, with the dimensions reversed at the tail. The remaining lines exercise
omitted and explicit `med` sizes, a 1 px stroke, rotation with horizontal flip,
absent ends, explicit `none`, and bent and curved preset geometry. The arrow row is dashed.

[ECMA-376 Part 1](https://ecma-international.org/publications-and-standards/standards/ecma-376/),
sections 20.1.10.32–34, defines the three relative sizes and illustrates the five
marks. It depicts `arrow` as an open chevron and `stealth` as a notched arrow.
It does not prescribe numerical size multipliers.

The renderer uses 2, 3, and 5 times `max(stroke width, 0.7 mm)` for `sm`, `med`,
and `lg`; omitted dimensions use `med`. The filled-marker scale, minimum base,
and stealth notch at 60% of the length follow
[LibreOffice's DrawingML importer](https://github.com/LibreOffice/core/blob/master/oox/source/drawingml/lineproperties.cxx).
Open-chevron dimensions describe the stroke centerline, so its stroke extends
beyond those dimensions; LibreOffice instead builds an outlined polygon.
Diamond and oval marks are centered on their path endpoints.

For example, `triangle-lg-med` is a blue (`#315EFB`) 4 px line from `(880, 120)`
to `(1140, 120)`. Its head measures 20 × 12 px and its tail 12 × 20 px. On main,
its stroke JSON is `{"color":"#315EFB","width":4.0}`. With line ends enabled,
only `headEnd` and `tailEnd` are added. Both plain lines keep exactly that JSON.

`line-ends-before.png` and `line-ends-after.png` replay slide 1's actual display
lists through the main and branch canvas implementations at 1× scale.

# Picture crop and mask

`picture-crop-mask.pptx` derives from the tracked demo deck. Only slide 1's picture and its bitmap change; slides 2 and 3 remain controls.

The 1024×1024 bitmap has RGB `(floor(x / 4), floor(y / 4), 40)`. The picture uses `srcRect l="10000" t="20000" r="30000" b="10000"`, an ellipse, and a 2px `#FF00FF` outline in the frame `(100, 50, 200, 100)` CSS pixels.

The kept source rectangle is `(102.4, 204.8, 614.4, 716.8)` pixels. Rendering should crop that rectangle, clip it to the ellipse, and stroke the ellipse. The adjacent before/after PNGs show the isolated picture on a 400×200 surface.

`text-baseline-script.pptx` derives from the same deck. Only slide 1's subtitle changes: one paragraph whose `a:pPr/a:defRPr` carries `baseline="0"` and whose runs alternate between no shift, `baseline="30000"` and `baseline="-25000"` at `sz="1700"`. The zero must stay a no-op; the other two must shrink and shift their runs.
