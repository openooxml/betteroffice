---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

The accessibility page mirror now updates in place after an edit: only the blocks whose content changed are rebuilt, and blocks that only moved get their document positions shifted. On large documents this removes most of the per-keystroke main-thread layout and style work.
