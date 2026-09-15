---
"@betteroffice/rust-crates": patch
---

Mask threaded-comment persons, pivot caches and tables, connections, and external-link details during XLSX redaction. Person display names and user IDs, threaded-comment text and timestamps, pivot field names with shared-item and record values (numbers keep the numeric placeholder, dates use the epoch, errors use `#N/A`), pivot-table and slicer names and captions, connection strings with commands and URLs, external sheet names with cached values, DDE service/topic/item names, and OLE program IDs are now masked, while shared-item indexes, relationship IDs, and person GUID links keep resolving.
