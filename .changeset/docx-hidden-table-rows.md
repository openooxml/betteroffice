---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': minor
---

DOCX table rows marked hidden (`w:trPr/w:hidden`) whose content is all hidden no longer paint borders or shading or take height unless hidden text is shown. Marked rows with visible content still render. A fully hidden table takes no space.
