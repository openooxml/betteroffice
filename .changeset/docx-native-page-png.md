---
'@betteroffice/docx-native': minor
---

The raw `EditSession.render_page_png(displayListJson, page, scale)` rasterizes one laid-out page to PNG in the native engine, from the same display list the canvas paints, with the session's fonts. Oversized pages and missing fonts are refused with stable messages.
