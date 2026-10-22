---
"@betteroffice/pptx-react": patch
---

An edit made just before the presentation worker fails now still reaches `onChange`, so hosts keep their unsaved-changes prompt after a crash.
