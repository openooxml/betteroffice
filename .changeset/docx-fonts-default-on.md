---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/agents": patch
---

DOCX measurement and browser font loading now use the bundled fonts manager by default, loaded lazily. Configured fonts and document embeds still win; families the manager lacks still load from Google Fonts unless disabled.
