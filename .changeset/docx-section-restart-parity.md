---
"@betteroffice/docx": patch
---

Odd page and even page section breaks into a section that restarts its page numbering now decide parity from the restarted number, as Word does, so no extra blank page is added when the restarted number already has the right parity.
