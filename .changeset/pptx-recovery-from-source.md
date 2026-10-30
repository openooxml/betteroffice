---
'@betteroffice/pptx': patch
'@betteroffice/pptx-react': patch
---

Crash recovery in the PowerPoint editor now works on picture-heavy decks: it reopens the original file and replays the edits instead of loading an open-time state export that could exceed the 64 MiB update limit, and opening no longer spends time exporting that state.
