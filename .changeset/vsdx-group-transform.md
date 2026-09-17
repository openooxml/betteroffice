---
"@betteroffice/vsdx": patch
"@betteroffice/rust-crates": patch
---

Keep a group's transform on its display-list node instead of baking it into every child, and derive that transform from the group's own cells rather than the live bounding box of its children. Editing or removing one child no longer moves or resizes its siblings, hit testing and line jumps compose ancestor transforms, and `bounds_affine` drops its child-extent argument.
