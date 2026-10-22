---
"@betteroffice/xlsx": minor
"@betteroffice/agents": minor
---

Adds conditional formatting: `conditionalFormats(sheet)` lists rules, and agents add or delete `cellIs` and `expression` rules with a differential format via `xlsx_list_conditional_formats` and `xlsx_edit_conditional_formats`. Structural edits that would invalidate session-authored rules are refused; source rules keep their geometry.
