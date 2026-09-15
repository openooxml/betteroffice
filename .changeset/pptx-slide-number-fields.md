---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Evaluate slide-number fields on masters and layouts, counting from the presentation's first slide number.

Collaboration snapshots use deck schema 2.1. Older snapshots migrate to deck
schema 2.1; missing starting numbers default to one. Readers supporting only
older schemas reject 2.1 snapshots.
