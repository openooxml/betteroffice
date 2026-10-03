---
"@betteroffice/docx-react": minor
"@betteroffice/docx": patch
---

The editor ref adds async twins for its synchronous document members (`readSelectionInfo`, `findParagraphs`, `scrollToParagraph`, `scrollToComment`, `scrollToChange`, `insertComment`, `insertCommentReply`, `onDocumentChange`) and deprecates the synchronous ones. In viewer sessions `getDocument`, `getPageContent` and `findInDocument` throw `DocxAsyncOnlyError`, and selections reach `onSelectionChange` and plugins.
