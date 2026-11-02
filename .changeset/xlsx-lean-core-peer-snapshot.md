---
"@betteroffice/xlsx": patch
---

The lean spreadsheet engine core used by the editor's worker no longer carries the peer snapshot code that only full-core edit peers use, making each editor open's WebAssembly download smaller. Recovery, collaboration and exports are unchanged.
