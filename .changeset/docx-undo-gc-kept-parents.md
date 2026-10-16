---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

Docx sessions with undo no longer crash after a peer deletes a story or table this replica edited and undid. These replicas now keep deleted content until reload instead of freeing parents the undo history still points into.
