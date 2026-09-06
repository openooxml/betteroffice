---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Evaluate slide-number fields on masters and layouts, counting from the presentation's first slide number.

Collaboration snapshots use deck schema v4. Older snapshots migrate through the
existing v3 connector migration before the v4 slide-number migration; missing
starting numbers default to one. Readers supporting only v3 reject v4 snapshots.
