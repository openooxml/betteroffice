---
"@betteroffice/pptx-react": patch
---

The in-thread editor's `selectText` accepts backward ranges and keeps their direction. Both editor and plugin selections reject offsets past the final caret or inside a surrogate pair, matching the worker editor.
