---
"@betteroffice/agents": patch
---

Uses native engines by default on Node, Bun and Deno for editing and PNG previews, with automatic WebAssembly fallback. Browsers keep WebAssembly; `@napi-rs/canvas` is now an optional peer needed only for fallback previews.
