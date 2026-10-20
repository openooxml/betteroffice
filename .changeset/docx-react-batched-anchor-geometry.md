---
"@betteroffice/docx-react": minor
---

Plugin geometry adds required `readAnchorGeometries(targets)`, `getPageRect(pageIndex)` and `isPageBuilt(pageIndex)` (hand-written geometry fakes need them), and answers reaching unbuilt pages list `unbuiltPages`. New anchor failure codes `pending` (worker-only; retry it like `layout-unavailable`) and `read-failed` break exhaustive switches.
