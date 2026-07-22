---
"@betteroffice/rust-crates": patch
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-react": patch
"@betteroffice/xlsx-i18n": patch
---

Add a Google Sheets-style toolbar to the XLSX editor backed by new engine
APIs for range styling, number formats, selection-format aggregation, format
painting, merge queries, and history state. Formatting is fully collaborative
through a content-addressed style catalog (collaboration schema v3; v2 state
does not migrate). Merging replaces intersecting ranges like Excel, parsing
repairs overlapping merges in third-party files, and display-list font fields
now serialize correctly so styled text renders with its real font, size, and
weight.
