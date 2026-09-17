---
"@betteroffice/vsdx": minor
"@betteroffice/vsdx-react": minor
"@betteroffice/vsdx-i18n": patch
---

Validate a diagram against five read-only rules: dangling connectors, isolated shapes, overlapping shapes, crossing connectors and empty required shape-data rows. `validate()` and `validatePage(pageIndex)` return deterministic, ordered issues addressed by session shape id, and `IssuesPanel` renders them.
