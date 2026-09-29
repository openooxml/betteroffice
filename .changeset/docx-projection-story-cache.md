---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Rebuild the editor's position projection incrementally: an edit now re-reads only changed paragraphs, via new `storiesChangedSince`, `storySegmentUnitDigests` and `storySegmentUnits` session queries.
