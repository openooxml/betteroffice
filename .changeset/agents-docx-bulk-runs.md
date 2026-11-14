---
'@betteroffice/agents': patch
---

Adds `docx_read_runs` and `set_runs` for paginated DOCX reads and text edits that preserve run formatting. Optional `maxEdits` and CLI `--max-edits` raise the default 32-edit cap up to 500.
