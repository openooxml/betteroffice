---
"@betteroffice/xlsx-react": minor
"@betteroffice/xlsx-i18n": patch
---

Adds an experimental `rowDecorations` plugin hook to the XLSX editor: plugins show badges and dots in a row gutter left of the grid, and a click or Enter sends `row-decoration-activate` to the owning plugin. Exhaustive `XlsxPluginEvent` handlers must handle this event.
