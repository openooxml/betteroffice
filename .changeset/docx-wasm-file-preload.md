---
"@betteroffice/docx": patch
---

`preloadOpcWasm`, `preloadParseWasm`, `preloadEditWasm` and `preloadLayoutWasm` now read a `file:` wasm asset from disk instead of handing its URL to `fetch`, matching what the synchronous path already did. This affects only hosts where the packaged asset resolves to a `file:` URL and the global `fetch` rejects that scheme: plain Node, and any Node or Bun process that has installed a DOM shim, since happy-dom replaces `fetch` with one that rejects `file:`. Browser and bundler consumers are unaffected — there the asset URL is http(s), so it still streams through `fetch` exactly as before, as does any other non-`file:` URL and any module or URL passed explicitly. Resolving the asset path also no longer passes a `URL` object to `fileURLToPath`, so a shim supplying its own non-native `URL` cannot defeat the disk read.
