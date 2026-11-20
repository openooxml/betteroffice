---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

A comment lost the range it was anchored to when the range sat inside a tracked change, a hyperlink or a content control and another paragraph was edited: the comment stayed, attached to nothing. A save now leaves such a paragraph as it was in the file and takes out only the markers of comments that were deleted.
