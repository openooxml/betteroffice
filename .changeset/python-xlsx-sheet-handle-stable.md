---
'@betteroffice/python-xlsx': patch
---

A held `Sheet` keeps addressing the same sheet as sheets are added, moved, or removed. A handle to a removed sheet raises `KeyError`.
