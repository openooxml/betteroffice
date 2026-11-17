---
'@betteroffice/docx-native': patch
'@betteroffice/pptx-native': patch
'@betteroffice/rust-crates': patch
---

Native rendering now draws JPEG and PNG pictures past the pixel cap at a reduced size, matching the canvas downscale, instead of leaving them out. Pictures whose reduced decode would still exceed the memory budget are refused as before.
