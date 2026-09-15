---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

Fall back to the main-thread engine when the resident worker crashes, times out, or answers corruptly mid-input so typing survives worker failures. Replay the pending keystroke on the main-thread engine and keep genuine engine-level input rejections surfacing as errors.
