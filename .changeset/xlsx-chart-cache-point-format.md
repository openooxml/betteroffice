---
'@betteroffice/xlsx': patch
'@betteroffice/xlsx-native': patch
---

Inserting or deleting rows or columns inside a range a chart plots through a numeric cache whose points carry their own number format now rebuilds the cache, with each point's format taken from its cell, instead of being refused.
