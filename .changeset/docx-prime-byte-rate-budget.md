---
'@betteroffice/docx': minor
---

Background save seed priming runs the committed scan in resumable slices. Each slice's byte budget follows the rate at which the last slice charged cost, growing at most twofold per slice; slices that charge nothing keep it. Saved bytes are unchanged.
