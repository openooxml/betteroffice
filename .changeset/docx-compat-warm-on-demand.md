---
"@betteroffice/docx-react": patch
---

Stop re-parsing the whole document on an idle callback after opening. The editor now builds its document model for `onChange` and `onContentChange` listeners once the first pages are on screen, and only when such a listener exists; Save, export and `getDocument()` build it on first use. On a 267-page document the idle re-parse ran before the first layout and delayed it by about 1.5 s, and it doubled the editing engine's memory.
