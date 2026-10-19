---
"@betteroffice/xlsx": minor
"@betteroffice/xlsx-react": patch
---

Fixes Ctrl/Meta+Arrow data edges, two-stage Ctrl/Meta+A selection, and navigation through hidden tracks, enabled by default in editors. `selectionKeyReducer` accepts optional `SelectionData` and `SelectionHidden`; `moveFocus` accepts `hidden`; `GridMeta` includes optional `hiddenRowSpans` and `hiddenColSpans`.
