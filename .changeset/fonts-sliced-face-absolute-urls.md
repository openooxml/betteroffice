---
'@betteroffice/fonts': patch
---

In apps bundled with Turbopack, CJK text in a document opened in the worker editor no longer stays blank: the font files it needs are now requested with full URLs the worker can load.
