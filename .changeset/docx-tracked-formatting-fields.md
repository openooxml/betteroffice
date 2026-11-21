---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
---

Rejecting a tracked formatting change keeps a field's result, and pasted tables get their own revision ids. A suggested hyperlink over a field result is refused with `SuggestedFormattingOverContainer`, and host `applyEdits` over a live formatting suggestion is refused as a tracked revision conflict.
