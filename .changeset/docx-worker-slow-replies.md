---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Keep editing on the resident engine worker when a large document takes over five seconds to reply, instead of falling back to the main thread; only a worker silent for a minute is replaced. Queued keystrokes and worker-failure recovery now apply in order.
