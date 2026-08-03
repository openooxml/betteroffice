---
"@betteroffice/pptx": patch
"@betteroffice/pptx-react": patch
"@betteroffice/rust-crates": patch
---

A deck snapshot persisted by the previous release opens again. Charts made the stored package a version 2 document, and the version check demanded an exact match, so every version 1 snapshot — every presentation a collaborator had already edited and saved — came back as `unsupported deck schema version` and could not be reopened.

`open_from_update` now migrates instead of refusing. A version 1 document hydrates, its stored package is read back and rewritten in the current shape, and the document is stamped version 2, so the next snapshot the session writes is a version 2 one and the upgrade happens once. Nothing else in the document changes: the slide order, slide, shape and story containers were already identical between the two versions, and the only difference was the chart list the stored package gained. That list is optional when reading, so a package written before charts existed loads with none rather than failing on a missing field. Two clients opening the same old snapshot write the same migration and converge.

A version this build does not know — a document from a newer release, or one whose version is missing or nonsense — is still rejected, and still reported before the stored package is parsed so the version is the error the caller sees.
