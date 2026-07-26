# betteroffice-pptx-parse

PresentationML parsing and part-preserving PPTX package writes.

`parse_pptx` reads a deck into the typed model: presentation properties,
slides, masters, layouts, shapes, text bodies, themes, and media, with the
relationship graph resolved. `parse_pptx_with_limits` takes explicit
`ParseLimits` for hostile input.

`write_pptx` re-zips the retained parts, so every part's bytes survive a round
trip unchanged and parts the engine does not model are carried through as-is.
The ZIP container itself is rebuilt, so the output is not byte-identical to the
input — do not rely on it for package hashes or signature validation.

Used by [betteroffice-pptx](https://crates.io/crates/betteroffice-pptx).

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
