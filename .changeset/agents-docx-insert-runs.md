---
'@betteroffice/agents': patch
---

`docx_edit` `insert_paragraphs` accepts optional `runs` (text with `bold`, `italic` and `underline`) in place of `text`, so mixed formatting needs no follow-up `format_text` edits. It works with `runFormat` and `list: continue`; `text` is unchanged.
