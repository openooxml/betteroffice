# Embedded font assets

## Carlito-Regular.ttf

- **Font:** Carlito Regular
- **Why:** metric-compatible with Calibri (Excel's default UI/body font), so text
  laid out against Calibri metrics renders at the same advances without shipping
  a proprietary font. Chosen for headless, deterministic rendering — the bytes
  are compiled into `xlsx-raster` via `include_bytes!`, so there is no system
  font access and output is identical on every machine and in CI.
- **License:** SIL Open Font License, Version 1.1 — see `OFL.txt` in this
  directory. The OFL applies to this font asset only, not to the crate's code
  dependencies (those are MIT/Apache, gated by `cargo-deny`).
- **Source:** https://github.com/google/fonts/blob/main/ofl/carlito/Carlito-Regular.ttf
  (fetched from `raw.githubusercontent.com/google/fonts/main/ofl/carlito/`).
- **Copyright:** Copyright 2013 The Carlito Project Authors
  (https://github.com/googlefonts/carlito), with Reserved Font Name "Carlito".

Glyphs missing from Carlito shape to glyph 0, which Carlito draws as a `.notdef`
tofu box. A real font-fallback chain (e.g. CJK coverage) is future work.
