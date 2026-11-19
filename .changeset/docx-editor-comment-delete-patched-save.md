---
'@betteroffice/docx': patch
'@betteroffice/docx-native': patch
'@betteroffice/rust-crates': patch
---

Deleting a comment in the editor and saving no longer forces a rewrite of every part. Where the writer can patch the body, a comment inside one paragraph leaves the other paragraphs' source bytes unchanged. Comment and bookmark ids differing only by leading zeros match.
