# betteroffice-pptx-parse

PresentationML parsing and byte-preserving PPTX package writes.

`parse_pptx` reads a deck into the typed model: presentation properties,
slides, masters, layouts, shapes, text bodies, themes, and media, with the
relationship graph resolved. `parse_pptx_with_limits` takes explicit
`ParseLimits` for hostile input.

`write_pptx` preserves the source package byte for byte, so parts the engine
does not own survive a round trip untouched.

Used by [betteroffice-pptx](https://crates.io/crates/betteroffice-pptx).

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
