---
"@betteroffice/desktop": patch
---

A failed document save now shows why it failed instead of "The editor is still opening the file", which appears only while the editor is still opening. Saves and room switches that overlap now serialize one at a time, each reporting its own failure.
