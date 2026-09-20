---
'@betteroffice/vsdx': patch
'@betteroffice/vsdx-react': patch
---

Add `probeCellWrites`, which asks the engine's mutation policy what a write to each cell would do without writing it, and have the editor gate its controls on that answer. The editor previously re-implemented the GUARD and SETATREF walk in TypeScript, so its idea of what was writable could drift from what the engine actually enforces; it now reads locks, GUARD interception and SETATREF redirection from the one policy that decides the write.
