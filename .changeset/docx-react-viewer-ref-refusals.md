---
'@betteroffice/docx-react': minor
---

With `experimentalWorkerOpen`, viewer sessions (read-only or viewing) answer synchronous edits without a main-thread copy: `getEditorRef` returns `null`, `setParagraphStyle`, `applyFormatting` and `insertBreak` return `false`, and comment insert and reply members return `null`. Tracked-change decisions and navigation are refused.
