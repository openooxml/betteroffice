---
"@betteroffice/fonts": patch
"@betteroffice/docx": patch
"@betteroffice/agents": patch
---

Uses Selawik for upright Segoe UI families by default, preserving their weights; italic Segoe UI keeps a real italic fallback face. Adds Office-aware `resolveGlyphFallback` and `resolveGlyphFallbackFace` coverage for characters absent from Selawik in Word and PowerPoint previews.
