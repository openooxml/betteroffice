---
"@betteroffice/docx": patch
---

`ResidentEngineWorkerClient` can open a DOCX in the resident worker (`open`, `fontRequirements`, `encodeState` and the `opened` bootstrap option), and `@betteroffice/docx/yrs` exports `preparedDocxDigest` and `decodeDocxHostJson`. Existing opens are unchanged.
