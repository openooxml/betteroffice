---
'@betteroffice/docx': patch
---

Paragraph offset docs now count leading block embeds: in a paragraph that follows one table, offset 0 is before the table and the text starts at offset 1; each further leading table or block content control adds one.
