---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
---

Rejecting a tracked formatting change keeps a field's result, and pasted tables get their own revision ids. A suggested hyperlink over a field result throws `SuggestedFormattingOverContainerError`, and a direct host replacement or deletion over a formatting suggestion is refused as a tracked revision conflict.
