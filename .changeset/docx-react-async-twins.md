---
"@betteroffice/docx-react": minor
"@betteroffice/docx": patch
---

The editor ref adds async twins for its synchronous document members (`readSelectionInfo`, `findParagraphs`, `scrollToParagraph`, `scrollToComment`, `scrollToChange`, `insertComment`, `insertCommentReply`, `onDocumentChange`) and deprecates the originals. Viewer sessions publish selections and throw `DocxAsyncOnlyError` from `getDocument`, `getPageContent` and `findInDocument`.
