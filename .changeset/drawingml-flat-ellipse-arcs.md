---
"@betteroffice/pptx": patch
"@betteroffice/docx": patch
---

Fixes cylinder (`can`) shapes with a near-zero top adjustment drawing their top ellipse slightly outside the shape frame. Arcs that start or end on an axis now stay exact on very flat ellipses.
