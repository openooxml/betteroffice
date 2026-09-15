---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Keep unmodeled inline drawings (OLE objects without image content, unreadable chart placements, drawing-owned compatibility wrappers) as their original markup through editor saves instead of dropping them, and save sessions whose chart runs predate drawing replay by restoring the placement from the source part or keeping the run out with a warning.
Tracked-change and hyperlink wrappers now survive around images, shapes, charts and opaque drawings on both seeders, OLE objects with an image fallback and picts whose image cannot be resolved are kept as opaque drawings, and anchor offsets count every embed as one unit on every layer, so bookmark and comment positions after an embed no longer shift by one across a save (parsed offsets after an embed move up by one accordingly). Save results carry `warnings` (dropped chart runs, opaque markup referencing a missing relationship), and seeding refuses a document whose opaque drawing payloads exceed 8 MiB with a typed error instead of exhausting memory.
