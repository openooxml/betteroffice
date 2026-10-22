---
"@betteroffice/xlsx": patch
---

Arithmetic on text holding a year-first date, such as `2024-03-05` or `2024/3/5 18:00`, reads its date serial as Excel does. Text padded with non-breaking or other non-ASCII spaces stays text and gives `#VALUE!`.
