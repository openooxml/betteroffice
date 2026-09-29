---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Type and delete in table cells through the resident engine worker, as in body paragraphs. A cell edit used to update the main-thread session, lay the whole document out on the main thread, and then send the worker a state diff to lay out again; it now reaches the worker as one request that lays out once, with the region pass re-measuring only the edited table.
