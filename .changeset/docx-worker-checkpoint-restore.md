---
"@betteroffice/docx": patch
---

The resident worker restores command checkpoints with caret, stored marks, selection dependencies, comments, layout reservations and proposal preview decisions. `openSource` restores the checkpoint's Yrs `state`; pass its `recoveryState` to `commandInit` alongside `version` and `throughSeq`.
