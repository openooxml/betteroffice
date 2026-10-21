---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
---

A `CollaborationProvider` on the worker editor's replica runs its sync protocol and presence in the document worker. The main thread forwards transport frames undecoded, one message per burst, and a burst of edits merges as one change. Its API and transport contract are unchanged.
