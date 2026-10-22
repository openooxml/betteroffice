---
"@betteroffice/docx": patch
"@betteroffice/pptx": patch
---

Saving or opening a very large document no longer ends the editing session after 60 seconds: the document worker reports progress while it works, so the watchdog only fires for a worker that has really stopped responding. Saves also skip a redundant package reparse.
