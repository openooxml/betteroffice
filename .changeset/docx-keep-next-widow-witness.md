---
'@betteroffice/docx': patch
'@betteroffice/rust-crates': patch
---

Weigh a `w:keepNext` group against the shortest slice of its follower that may legally start a page, not always the follower's first line. Placement already refuses to strand fewer than two lines of a paragraph across a page break, so a heading bound to a three-line paragraph could satisfy the keep test against one line, then watch widow/orphan control move all three to the next page and leave the heading behind. Both now read one predicate, so the keep witness is the whole follower when it cannot split, two lines when widow control governs it, and one line when the follower turns widow control off.
