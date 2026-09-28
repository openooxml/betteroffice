---
"@betteroffice/vsdx": minor
"@betteroffice/vsdx-react": patch
---

Render Visio linear fill gradients. The renderer resolves the `FillGradient` section through masters and styles, evaluates `GradientStopColor` with the theme so `THEMEVAL` stops pick up the document theme, and emits a gradient paint carrying `FillGradientAngle`. The display-list contract version moves from 4 to 5 and `Paint` gains an optional `angleDeg`.
