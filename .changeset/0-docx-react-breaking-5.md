---
"@betteroffice/docx-react": minor
---

**Breaking:** For plugin authors: hand-written geometry objects must implement `readAnchorGeometries`, `getPageRect` and `isPageBuilt`, and anchor failures add the `pending` and `read-failed` codes. Update test fakes and exhaustive switches over failure codes.
