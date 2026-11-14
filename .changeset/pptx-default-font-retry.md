---
'@betteroffice/pptx': patch
---

Default fonts that load after the first paint's five-second wait now register and re-layout the deck instead of being dropped for the session. Failed loads retry up to three times, after two and eight seconds.
