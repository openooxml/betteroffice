---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Reads now include `drawings` on content-control and field atoms by default, including cached field results. Drawing geometry targets accept `index` for control drawings; cached field-result drawings return `missing-target` because they have no individual display position.
