---
"@betteroffice/agents": patch
"@betteroffice/docx": patch
"@betteroffice/docx-i18n": patch
"@betteroffice/docx-react": patch
"@betteroffice/fonts": patch
"@betteroffice/fonts-cjk": patch
"@betteroffice/fonts-intl": patch
"@betteroffice/fonts-latin": patch
"@betteroffice/pptx": patch
"@betteroffice/pptx-i18n": patch
"@betteroffice/pptx-react": patch
"@betteroffice/vsdx": patch
"@betteroffice/vsdx-i18n": patch
"@betteroffice/vsdx-react": patch
"@betteroffice/xlsx": patch
"@betteroffice/xlsx-i18n": patch
"@betteroffice/xlsx-react": patch
---

Package exports now carry a `default` condition next to `import`, so webpack consumers (`next build --webpack`, including editors loaded with `next/dynamic`) resolve every entry instead of failing with "Package path is not exported".
