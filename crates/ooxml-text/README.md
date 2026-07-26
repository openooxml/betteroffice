# betteroffice-ooxml-text

Text shaping and measurement for the OOXML layout engines. Turns run text into
positioned glyphs and line-break decisions with no browser APIs in the loop.

- `font_store` — registry over raw font **bytes**, never font names, with
  `head`/`hhea`/`OS/2` metrics, cmap lookup, advance widths, and an ordered
  fallback-chain resolver
- `shape` — OpenType shaping through rustybuzz, returning cluster-mapped glyphs
  with advances and offsets scaled to the requested size
- `line_break` — UAX-14 break opportunities, surrogate-safe and CJK-aware
- `bidi` — paragraph-level Unicode Bidirectional Algorithm runs
- `word_metrics` — the Word-specific rules: single-spacing line boxes from OS/2
  win metrics, auto/exact/atLeast line rules, justification gating and space
  stretch, the `w:kern` threshold, and the `settings.xml` compat flags that feed
  them
- `outline` — glyph outline extraction as font-unit path commands, from the same
  skrifa bytes the metrics came from

Callers hand this crate font bytes plus a fallback chain of `FontId`s.
Resolving a `w:rFonts` name to bytes — embedded `.odttf`, bundled
metric-compatible faces, Local Font Access, browser-measured fallback — stays on
the host side, which keeps results deterministic and identical across web and
native shells.

No `wasm-bindgen` here by design; a thin facade can wrap it.

Snap-to-grid (`w:docGrid`) is not implemented.

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
