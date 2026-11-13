---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
---

DOCX tables now show the cell shading and borders of their table style's conditional regions, such as header rows and banded rows. A `tblLook` with named attributes is read from those alone; its hex value applies only without them. Saving keeps these as style formatting.
