---
"@betteroffice/docx": patch
---

Saving formatting changes in DOCX paragraphs with many short runs is faster: each run's text is looked up through an index instead of scanning the whole paragraph for every run.
