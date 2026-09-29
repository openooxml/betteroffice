---
"@betteroffice/docx": patch
---

Keep no frame reply alive through a retained page. A retained page's primitive-id array was a view into the whole FrameDelta buffer it arrived in, so one page kept a multi-megabyte reply in memory for as long as it stayed unchanged; retained pages now own a copy of their ids. The resident worker no longer copies every page's ids on each reply to work around this.
