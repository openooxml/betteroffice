---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Place absolutely positioned VML drawings at their authored offsets. A `w:pict`/`w:object` picture recorded its `margin-left`/`margin-top` in a field no renderer read, so every such drawing collapsed onto its anchor band's origin — a header picture landed at the top-left of the page instead of its offset, and the header it belongs to measured short, lifting the body text with it. Group-child `left`/`top` coordinates stay out of the anchor, since they position a shape inside its group rather than on the page.
