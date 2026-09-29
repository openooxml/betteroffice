---
"@betteroffice/docx-react": patch
---

Keep the whole-document position projection out of the moment a document opens. The first selection event of a session now waits for idle time, and the selection overlay, the selected-image check and the floating comment button no longer map the selection before there is a display list to draw it with. On a 267-page document this takes about 200 ms of main-thread work out of the commit that mounts the editor.
