---
"@betteroffice/docx": minor
"@betteroffice/docx-react": minor
"@betteroffice/agents": patch
---

DOCX measurement and browser font loading now use the bundled fonts manager by default, loaded lazily. Configured fonts and document embeds still win; families the manager lacks still load from Google Fonts unless disabled.
