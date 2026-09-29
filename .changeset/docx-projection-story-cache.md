---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Rebuild the editor's position projection from what changed. The session now stamps each story with the revision it last changed at (`storiesChangedSince`) and digests a story's segments paragraph by paragraph (`storySegmentUnitDigests`, `storySegmentUnits`). The paged editor keeps story segments across projections: an edit re-reads only the changed paragraphs of the changed stories instead of every segment of every story. Paragraph digests of stories read whole are fetched while the main thread is idle.
