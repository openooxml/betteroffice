---
"@betteroffice/pptx-react": minor
"@betteroffice/pptx": minor
---

Editable `PptxEditor` sessions with `experimentalWorkerOpen` now use the worker-owned editor, with `saveAsync`, `handleAsync` and `recoverySave` on the editor API. The flag remains experimental and opt-in, the default editor is unchanged, and collaboration is not supported in this mode yet and throws `PptxWorkerEditorCollaborationError`.
