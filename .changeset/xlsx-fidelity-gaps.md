---
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
"@betteroffice/rust-crates": patch
---

Formulas referencing defined names now resolve correctly, frozen panes render, and hyperlinks survive the round trip. The collaboration schema advances to version 5 and upgrades version 3 and 4 snapshots when read, so a client on this release cannot share a collaboration room with an older one: upgrade every peer together.
