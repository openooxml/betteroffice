---
'@betteroffice/pptx-native': patch
---

Now exposes exactly the `@betteroffice/pptx` API (same entry points, exports and types) on the native engine for Node.js, Bun and Deno; pick the engine by import. The `browser` condition throws. The `openPresentation(Buffer)`/`Presentation` facade is removed.
