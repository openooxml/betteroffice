---
"@betteroffice/docx-react": patch
"@betteroffice/docx": patch
---

The default worker editor's ref now serves `scrollToComment`, `scrollToChange`, `insertCommentReply`, `getComments`, `resolveComment` and `loadDocument`. Replies anchor where their parent is. Worker comment updates, range changes and removals report every comment in `editState.commentChange`.
