---
'@betteroffice/rust-crates': minor
---

`betteroffice-docx` orders font fallbacks as the product does: a family's own face, the Word last-resort face, the glyph fallback, then script faces. Latin text without a bundled face no longer gets a CJK face and too-tall lines. Adds hidden `FontManager::resolve_with_glyph_fallback`, naming the glyph fallback entry.
