---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

An inline SmartArt stays in its paragraph's flow, between the text around it, instead of being laid out above the paragraph. Saving an edit after a page or column break that follows the diagram keeps the break after it.
