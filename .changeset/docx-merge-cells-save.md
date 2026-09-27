---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Keep the text of merged table cells when a document is saved or projected again after a merge; an earlier save, `getDocument()` or `onChange` projection no longer makes later saves drop the merged-in cells' content or later edits to a table the merge moved into the surviving cell. Table receipts gain `changedStoryIds`, the existing stories an operation rewrote.
