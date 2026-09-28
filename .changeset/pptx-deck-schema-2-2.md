---
'@betteroffice/pptx': minor
'@betteroffice/pptx-react': minor
'@betteroffice/python-pptx': minor
'@betteroffice/rust-crates': minor
---

**Breaking:** collaboration updates now carry deck schema 2.2, and a 0.1.x update is migrated to 2.2 when it is opened. 0.1.x clients cannot open a document this version saved or migrated, so mixed-version collaboration across this upgrade is not supported: upgrade every client that shares a stored update.
