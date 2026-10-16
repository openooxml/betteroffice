---
"@betteroffice/docx": patch
"@betteroffice/rust-crates": patch
---

The delete-only `'local'` update a replica sends after a remote update now carries only deletions the replica made itself, so a peer that has not seen a concurrent winning edit keeps the losing edit's paragraph attributes and ids.
