---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Fix the layout collapse on documents dense with tracked changes. When no measurement face is registered — the default for a host that does not supply `measurementFontProvider` — the measure engine refuses every paragraph and the layout falls back to an estimated extent. That estimate put the whole paragraph on one line and then billed the line at most the column width, so a paragraph wider than its column had every run squeezed into a fraction of its own estimated advance: consecutive runs painted on top of one another and the line ran off the page. A tracked deletion and its replacement insertion sit in one paragraph and roughly double its text, which is why documents where most paragraphs carry a revision collapsed while the same file laid out cleanly once the changes were accepted.

The fallback now fills lines to the column width and reports each line's own estimated width, so every run is allotted the space its text needs and long paragraphs wrap instead of overprinting.
