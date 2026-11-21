---
'@betteroffice/desktop': patch
---

In the desktop app on macOS, a document could stop responding with "Worker commands are not initialized" once its editor loaded extra code, such as fonts for CJK text. Each editor worker is now built as one file.
