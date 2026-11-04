---
"@betteroffice/xlsx-react": patch
"@betteroffice/xlsx-i18n": patch
---

Adds experimental `cellEditable` and `beforeCellCommit` plugin hooks: plugins can lock cells for typing and clearing, rewrite typed input, or refuse it with `host-refused`. Works in the in-thread and default worker editors; the legacy worker editor ends plugins that define them.
