---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

When collaborators move or restyle the same shape at once, the replica whose change loses no longer forwards a deletion that left other replicas with a shape missing its position, which made them refuse the update.
