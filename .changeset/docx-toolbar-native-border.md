---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

DOCX toolbar icon buttons no longer inherit the browser's default `2px outset` border or ButtonFace fill in hosts that do not load Tailwind preflight. Explicit `border` and `bg-*` utilities still win.
