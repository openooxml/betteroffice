---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Opening a Word document no longer downloads the save engine up front. The editor prepares saving as soon as you edit or reach for Save, so the first save stays fast and read-only opens download less.
