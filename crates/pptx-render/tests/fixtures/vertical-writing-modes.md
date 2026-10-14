# Vertical writing modes

`vertical-writing-modes.pptx` is a synthetic one-slide repro for the `a:bodyPr/@vert` modes that
`vert`/`vert270` support left behind. It carries no private presentation content; its master,
layout and theme come from `text-orientation.pptx` with the placeholder `vert` attributes removed,
so nothing inherits a direction.

Six 1600200 × 3429000 EMU boxes at y = 1600200, captioned with their mode:

- Shape 40 `eaVert` and shape 44 `vert` hold the same sentence at 16 pt. Both must turn 90° and
  lay out in the swapped 3429000 × 1600200 box, wrapping to two columns that advance right to
  left. Their lines must land on the same coordinates.
- Shape 41 `mongolianVert` holds that sentence too. It turns the same way, but its columns advance
  left to right, so its lines occupy the same span in reverse order.
- Shapes 42 `wordArtVert` and 43 `wordArtVertRtl` hold `STACK` at 20 pt, centred. The box does not
  turn; each character becomes its own line, upright, in a single column.
- Shape 45 is a horizontal control.

[ECMA-376 Part 1](https://ecma-international.org/publications-and-standards/standards/ecma-376/),
`ST_TextVerticalType` (§20.1.10.83), defines `eaVert` as vertical text where East Asian glyphs stay
upright and other scripts turn, `mongolianVert` as `eaVert` with the flow going left to right, and
`wordArtVert` as "one letter on top of another".

LibreOffice renders `eaVert` exactly like `vert` and `wordArtVert` as an upright stack, which is
what this fixture asserts. It renders `mongolianVert` and `wordArtVertRtl` horizontally; PowerPoint
does not, and neither does this renderer.

Latin runs are laid out with the glyphs turned in `eaVert` and `mongolianVert`, as `vert` does.
Keeping East Asian glyphs upright inside a turned line is not implemented, and neither is the
column wrap that PowerPoint applies when a `wordArtVert` stack is taller than its box: the stack
stays one column and reports `overflow`. `mongolianVert` reverses its columns inside the block the
anchor produced rather than restarting them at the opposite edge of the box, so a partly filled box
anchored `t` or `b` keeps the block on the side `vert` would put it; the fixture anchors centred,
where the two agree.
