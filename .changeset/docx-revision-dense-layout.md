---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Fix the layout collapse on documents dense with tracked changes. When no measurement face is registered — the default for a host that does not supply `measurementFontProvider` — the measure engine refuses every paragraph and the layout falls back to an estimated extent. That estimate clamped the paragraph to the column width, so consecutive revision runs received slots narrower than their text and painted on top of one another.

The fallback now keeps the paragraph on one line and gives every run a conservative, script-neutral slot. Browser paint from this no-face path is clipped to its slot, so unusually wide glyphs are boundedly truncated instead of overprinting the following run. Paragraphs that need real font metrics continue to overflow rather than relying on guessed line breaks; registered-face text and its visual effects remain unclipped.
