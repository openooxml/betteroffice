`table-rotated.pptx` is `table-basic.pptx` with `rot="5400000"` on the graphic
frame's `p:xfrm`, and nothing else changed.

A rotation or flip pivots around the centre of the bounds the primitive carries,
so a table that grows past its frame would turn about the grown box rather than
the frame. The fixture pins that a turned table keeps its frame height and clips
instead of moving. No corpus deck rotates a table, so this case exists only here.
