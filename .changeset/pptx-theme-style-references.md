---
"@betteroffice/pptx": patch
"@betteroffice/rust-crates": patch
---

Resolve PPTX theme fill and line references, including background fills and placeholder colour transforms. Preserve explicit shape and placeholder formatting. Preserve font reference colours from the existing text-style resolver.

Migrate v1–v4 deck snapshots to schema 5 after the existing connector and slide-number migrations, preserving edits, numbering and source ordinals.
