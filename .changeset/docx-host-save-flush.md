---
"@betteroffice/docx-react": minor
---

Add `onSaveRequest` for host-controlled File > Save and Cmd/Ctrl+S, plus awaited `flushPendingInput()` on the editor refs. Built-in saves flush pending input before serialization, and concurrent UI save requests share one workflow.
