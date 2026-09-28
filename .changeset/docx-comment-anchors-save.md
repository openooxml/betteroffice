---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Save a comment moved with `setCommentRanges()` or a raw `setComment`, also into another story or by two collaborators at once, with one range and one reference at its new place, through undo, redo and collaborating replicas, instead of keeping the old reference beside a new one. Every save writes one reference in each paragraph that ends a comment's range, whatever order collaborators' changes arrived in; a reference that sits in another paragraph than its comment's range end moves after that end.

Save comment ranges that start or end inside a hyperlink, including one that holds a field or an equation, or inside a tracked change, or that follow a field, content control, equation, table or page break in their paragraph, as a paired start and end with a reference over the same text, and reopen them over that text; the editor's Save now writes the ranges of comments added in the editor.
