# Third-party notices

This wheel embeds third-party material in the compiled extension module. The
package itself is Apache-2.0; the notices below apply to the bundled material.

---

## Carlito (metric-compatible with Calibri)

`Carlito-Regular.ttf` is vendored unmodified from the google/fonts repository,
`ofl/carlito` (https://github.com/google/fonts); upstream project
https://github.com/googlefonts/carlito. The bytes are compiled into the
`betteroffice-xlsx-raster` crate via `include_bytes!` and are therefore present
inside this wheel's extension module. The renderer uses the face to measure and
draw cell text.

Copyright 2013 The Carlito Project Authors, with Reserved Font Name "Carlito".

License: SIL Open Font License, Version 1.1 — the full text is distributed with
this package as `licenses/Carlito-OFL.txt`. The OFL applies to the font asset
only, not to this package's code.
