---
"@betteroffice/agents": patch
---

PPTX reads distinguish soft line breaks (`\u000b`) from paragraphs (`\n`). `replace_text` keeps matching run formatting, including when removing trailing soft breaks, and otherwise uses the paragraph's first run format. Use `find_replace` or `office_propose` for partial edits.
