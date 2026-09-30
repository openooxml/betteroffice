---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

The resident worker now finishes a document's first layout in short steps, so requests that arrive meanwhile run between them, with identical results. `YrsSession.beginRegionLayout` and `resumeRegionLayout` expose the stepped region layout.
