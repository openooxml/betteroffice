---
"@betteroffice/docx": patch
---

Reads no longer refuse documents with tracked paragraph splits and merges. They show pending paragraph-mark revisions as accepted in the accepted view or as rejected in the original view, keeping the following paragraph's id when paragraphs join.
