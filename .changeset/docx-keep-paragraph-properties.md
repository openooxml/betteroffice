---
"@betteroffice/docx": patch
---

Saving a rebuilt paragraph now keeps its authored `adjustRightInd` and `cnfStyle` properties, and style inherited `autoSpaceDE` and `autoSpaceDN` values stay inherited instead of becoming direct formatting on numbered paragraphs.
