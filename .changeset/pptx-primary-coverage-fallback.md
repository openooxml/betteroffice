---
"@betteroffice/pptx": patch
---

When browser coverage fallback is enabled, text missing from its requested font now uses another registered font that covers it before reaching browser fallback. The selected font supplies both layout advances and glyphs.
