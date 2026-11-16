---
'@betteroffice/desktop': patch
---

Keep desktop editor web workers self-contained by disabling worker code splitting so WebKit does not re-evaluate worker entry modules across dynamically loaded chunks.
