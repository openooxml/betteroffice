---
"@betteroffice/pptx": patch
"@betteroffice/pptx-react": patch
---

Moves active slide painting into the editor worker by default. Slides opened before the worker is ready paint on the main thread, and the main painter remains the fallback for unsupported environments and images.
