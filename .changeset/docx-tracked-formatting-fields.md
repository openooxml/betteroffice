---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
---

Rejecting a tracked formatting change keeps a field's result, and pasted tables get their own revision ids. A suggested hyperlink over a field result throws `SuggestedFormattingOverContainerError`. Direct host replace, delete, `formatText`, `setParagraphStyle` and `setContentControlText` edits over a formatting suggestion are refused as tracked revision conflicts.
