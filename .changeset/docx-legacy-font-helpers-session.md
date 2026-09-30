---
"@betteroffice/docx": patch
---

The legacy module-level font helpers (`clear_measure_fonts`, `register_measure_font`, `measure_paragraph_json`, `outline_glyph_json`, `loadGlyphOutlineProvider`) must not be called while an editor session is open; use the session's methods.
