# Metafile gradients and pattern blits

`metafile-gradient.pptx` is a synthetic, publishable reproduction of the EMF
records the replayer refused. Its single 240 x 240 px picture embeds a
hand-built 200 x 200 unit EMF.

The top half is one `GRADIENTFILL` in `GRADIENT_FILL_TRIANGLE` mode: four
vertices and two triangles, `4FA9E0` along the top edge and `2B4E9B` across the
middle. The lower left quadrant is a second `GRADIENTFILL` in
`GRADIENT_FILL_RECT_H` mode, `F2E4C2` to `217A3C`. The lower right carries two
`BITBLT` records that name no source bitmap: `PATCOPY` fills (100,100)-(200,160)
with the selected `E03A2F` brush, and `DSTCOPY` over (100,160)-(200,200) must
leave the slide untouched. See [MS-EMF, sections 2.3.5.12 and
2.3.1.2](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-emf/).

Each gradient replays as flat-shaded bands, one display-list shape per band:
64 for each triangle and 64 for the rectangle, plus one shape for the pattern
blit, so slide 1 holds 193 shapes and no image. Bands are sampled at their
midpoint, so the outermost band sits half a band inside its vertex colour
(`2b4f9c` and `4fa8df` for the triangle pair). Each band spans from its own edge
to the far end of the shape, so neighbouring bands overlap and no antialiased
crack shows between them.

`metafile-gradient.before.png` was rendered on main `0f474aa5`, where an
unsupported record rejects the whole metafile and the picture is skipped; the
after image uses this implementation. Both are the picture box with a 16 px
margin, at scale 2.

LibreOffice's EMF import draws the rectangle gradient and the pattern blit but
leaves triangle-mode gradients blank, so its render of this file is not a full
expectation for the top half.
