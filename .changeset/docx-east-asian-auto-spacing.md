---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-docx': patch
---

Insert Word's East Asian auto-space (`w:autoSpaceDE`, `w:autoSpaceDN`, both default on) where East Asian text meets Latin letters or digits. The gap is a quarter em of the East Asian side, measured off Word's own PDF exports: an East Asian character standing before a Latin one advances 1.250 em against 1.000 em before another East Asian character, and a document that switches the feature off measures 1.000 em on both. Nothing is inserted next to a space, the two settings gate their own boundary, and both are parsed, honoured and written back.
