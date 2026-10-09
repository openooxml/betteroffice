---
"@betteroffice/docx": patch
---

Run font sizes outside Word's 1–1638 pt range, such as `w:sz="0"`, now lay out and paint at the nearest size in that range, so paged structured exports of such documents succeed instead of returning `layout-unavailable`.
