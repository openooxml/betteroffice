---
'@betteroffice/docx': patch
---

Word text watermarks now draw on the pages whose header carries them, stretched to their shape box, and no longer on a first page whose header has none. `TextWatermark` gains optional `widthEmu` and `heightEmu`.
