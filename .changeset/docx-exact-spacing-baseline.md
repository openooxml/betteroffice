---
"@betteroffice/rust-crates": patch
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Paragraphs with `lineRule="exact"` or `lineRule="atLeast"` now put their text where Word puts it.

Two defects compounded here. Line measurement applied the paragraph's spacing rule, kept the ruled *height*, then discarded the ruled ascent and descent in favor of the raw font metrics — so a row could claim more ascent plus descent than it had height at all, and the painter placed that baseline below the bottom of its own line box, overlapping the row beneath and disagreeing with hit-test geometry. Underneath that, the spacing rule itself was wrong about both fixed rules.

Word's actual behavior was measured rather than inferred: probe documents rendered by Word 16.112, true glyph baselines read back from the PDF, each probe at the top margin of its own page so the line-box top needs no reference line. An `exact` box puts its baseline at a flat 80% of the box — a constant that depends on neither the font nor the size, confirmed by four families whose win-metric ratios span 0.780 to 0.810 producing byte-identical baselines at 12, 24 and 36pt. A floored `atLeast` box puts its slack *above* the ascent, preserving the content descent from the bottom. The engine previously had these backwards: it gave `exact` the descent-preserving treatment that belongs to `atLeast`, and gave `atLeast` leading below the descent that belongs to neither.

Both rules now match the measurement, the ruled ascent and descent reach the emitted row in the wrapped and empty-paragraph paths alike, and `ascent + descent <= lineHeight` is pinned by test across `exact` above and below the content height, floored and unfloored `atLeast`, sub-single `auto`, default spacing and image-bearing lines. Neither corrected rule leaves any leading, so the corrected rows land on the same baseline under either the half-leading model the painter uses or the leading-below-descent model the metrics document — which also removes a 16px error on `atLeast` text without changing the painter. `auto` spacing is untouched.
