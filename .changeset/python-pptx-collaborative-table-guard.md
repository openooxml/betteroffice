---
'@betteroffice/python-pptx': patch
'@betteroffice/rust-crates': patch
---

Presentations opened with `open_collaborative`, `open_collaborative_with_limits` or `open_from_update` now refuse table row and column insertions and deletions, as shared TypeScript sessions do. Peers can still edit cell text and merge those changes.
