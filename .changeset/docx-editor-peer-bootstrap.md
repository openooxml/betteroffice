---
"@betteroffice/docx": minor
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Worker-open editors build their editing copy from the worker's opened document instead of parsing the file a second time. The `yrs` entry adds `encodePeerMetadata`, `bootstrapPeer`, `PeerMetadataError` and `peerMetadataTags`.
