---
"@betteroffice/docx-react": minor
"@betteroffice/docx": patch
---

The editor ref adds async twins for its synchronous document members (`readSelectionInfo`, `findParagraphs`, `scrollToParagraph`, `scrollToComment`, `scrollToChange`, `insertComment`, `insertCommentReply`, `onDocumentChange`) and deprecates the originals. In viewer sessions selections reach `onSelectionChange` and plugins, and `getDocument`, `getPageContent` and `findInDocument` throw `DocxAsyncOnlyError`, which `DocxReplicaNotReadyError` retry loops do not catch.
