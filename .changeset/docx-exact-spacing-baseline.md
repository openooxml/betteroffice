---
"@betteroffice/rust-crates": patch
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Text in a paragraph with `lineRule="exact"` or sub-single line spacing now sits inside its own line box instead of hanging below it.

Line measurement applied the paragraph's spacing rule, kept the ruled *height*, and then threw the ruled ascent and descent away in favor of the raw font metrics. That is harmless for `atLeast` and for `auto` at or above single spacing, where the rule only adds leading below the descent and leaves ascent and descent alone. It is not harmless for the two rules that legitimately move them: `exact` fixes the box and eats the ascent side, and sub-single `auto` shrinks ascent and descent proportionally. A row could therefore claim more ascent plus descent than it had height at all — at 12pt, `exact 10` reported an ascent of 14.48 in a 10px box — and the painter placed that baseline up to 5px *below* the bottom of the line, so the text overlapped the row beneath it and paint and hit-test geometry disagreed with the document.

The ruled ascent and descent now reach the emitted row, in the wrapped path and in the empty-paragraph path alike, and `ascent + descent <= lineHeight` is pinned by test across `exact` above and below the content height, sub-single `auto`, `atLeast`, default spacing and image-bearing lines. Both rules that this corrects produce no leading, so the corrected rows place the baseline identically under either the CSS half-leading model the painter uses or Word's leading-below-descent model — the fix moves `exact` and sub-single text only, and leaves every other line where it was.
