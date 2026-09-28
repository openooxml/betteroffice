---
'@betteroffice/pptx': patch
'@betteroffice/rust-crates': patch
'@betteroffice/python-pptx': patch
---

Free a proposal's cached preview once it is rejected or accepted, so a long review session no longer grows memory with every proposal, and previewing a rejected or accepted proposal fails with "no proposal" again instead of returning its stale preview.
