---
"@betteroffice/rust-crates": patch
---

Harden single-line metric inputs while retaining quantization and typographic-metric variants as off-by-default experiments.

Direct AppleScript and PDF-baseline measurements of Word 16.112 rejected whole-pixel metrics in compatibility modes 14 and 15. Float line pitch matched within ±0.019pt across four fonts, seven sizes, and 20-line spans; per-component quantization missed by up to ±0.49pt, with sign varying by size. Float advances matched within ±0.0145pt across 10–40-character runs; whole-pixel advances missed by up to 0.348pt.

`CompatFlags::gdi_line_metrics` and `CompatFlags::typo_line_spacing` therefore remain independent, opt-in measurement switches. Neither is enabled by paragraph input. Their bounded path preserves signed negative leading through paragraph aggregation, and `Auto(240)` preserves such a box exactly.

For finite sizes through 1638pt, the default `single_line_box` is bit-for-bit unchanged for every spec-valid metric. Separately, `Auto(240)` now returns its input exactly instead of introducing an f32 reassociation difference of up to 0.0000076px in the measured 72pt-and-under sample, or 0.00012px at the cap.

Input hardening rejects non-finite sizes and caps every line-box measurement at Word's 1638pt limit. Glyph advances remain uncapped and scale at the requested size, so larger requests produce line boxes and advances at different scales. Font-component bounds apply only after an experiment is enabled.
