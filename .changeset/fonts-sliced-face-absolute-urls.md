---
'@betteroffice/fonts': patch
---

Sliced CJK faces load in a worker editor when the app is bundled with Turbopack. The file URLs sent to the worker are now absolute; Turbopack names bundled assets by path only, and the worker could not resolve the path, so CJK text stayed blank in a read-only DOCX.
