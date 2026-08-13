---
"@betteroffice/rust-crates": patch
---

Harden single-line metric inputs while retaining quantization and typographic-metric variants as disabled experiments.

Direct AppleScript and PDF-baseline measurements of Word 16.112 rejected whole-pixel metrics in compatibility modes 14 and 15. Float line pitch matched within ±0.019pt across four fonts, seven sizes, and 20-line spans; per-component quantization missed by up to ±0.49pt, with sign varying by size. Float advances matched within ±0.0145pt across 10–40-character runs; whole-pixel advances missed by up to 0.348pt.

`CompatFlags::gdi_line_metrics` and `CompatFlags::typo_line_spacing` therefore remain independent, opt-in measurement switches. Neither is enabled by paragraph input. Their bounded path preserves signed negative leading through paragraph aggregation, and `Auto(240)` preserves such a box exactly.

The shipping float path does not clamp font components, so every spec-valid metric retains its prior result. Input hardening still rejects non-finite sizes and caps direct callers at Word's 1638pt limit; component bounds apply only after an experiment is enabled.
