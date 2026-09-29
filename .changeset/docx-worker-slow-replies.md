---
"@betteroffice/docx": patch
"@betteroffice/docx-react": patch
"@betteroffice/rust-crates": patch
---

Keep the resident engine worker when a large document takes longer than five seconds to answer a keystroke, instead of moving editing to the main thread for the rest of the session; only a worker that stays silent for a minute is replaced. Text typed while the worker is busy still joins one resident request, now without jumping ahead of a queued Backspace, Enter or arrow key, and Backspace or Delete presses queued behind busy input join one resident deletion laid out once. After a worker failure the host engine's recovery frame is numbered after the last worker frame, so it is applied instead of rejected as stale.
