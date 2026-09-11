# Floating header pagination

The regression fixture is generated in memory by `tests/support/quality_fixture.rs` and exercised by `cargo test -p betteroffice-docx-edit --test anchored_header`. Its twelve paragraphs use exact 18-point line spacing. A 50-pixel-wide header decoration is anchored at page coordinates (1010, 56.667) with `wp:wrapNone`.

Before the fix, lowering discarded the anchor. The header's 581.533-pixel extent expanded the top margin and forced one paragraph onto each of twelve pages; the header shape itself was not painted. After the fix, the body fits on one page at its authored top margin, and the decoration paints on the right.

Local before/after PNGs show page one at 150 DPI using the same CDN fonts. They are renderer comparisons, not Microsoft Word reference images, and are not committed.
