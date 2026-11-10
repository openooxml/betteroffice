---
'@betteroffice/docx': patch
'@betteroffice/agents': patch
'@betteroffice/rust-crates': minor
---

PNG export from a session bootstrapped from another, such as an agent preview, draws EMF and WMF pictures as the seeded session does, instead of dropping their text and bitmaps on the native engine.
