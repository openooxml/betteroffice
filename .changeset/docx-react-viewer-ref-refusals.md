---
'@betteroffice/docx-react': minor
---

With `experimentalWorkerOpen`, read-only and viewing editors answer synchronous edits without loading the document on the main thread: `getEditorRef` returns `null`, `setParagraphStyle`, `applyFormatting` and `insertBreak` return `false`, and `addComment`, `replyToComment`, `insertComment` and `insertCommentReply` return `null`. Accepting, rejecting and stepping through tracked changes are refused in these sessions.
