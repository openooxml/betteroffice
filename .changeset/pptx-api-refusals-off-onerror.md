---
"@betteroffice/pptx-react": patch
---

Editor API calls that are refused, such as `propose()` with an unsupported edit, now report the refusal only to the caller and no longer also call `onError`. Worker failures still call `onError`.
