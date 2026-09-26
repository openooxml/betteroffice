# @betteroffice/python-docx

## 0.2.0

### Minor Changes

- 67e0ae4: Release the accumulated DOCX, PPTX, and XLSX engine improvements in the Python bindings as minor updates.

### Patch Changes

- 59d71c4: Insert Word's East Asian auto-space (`w:autoSpaceDE`, `w:autoSpaceDN`, both default on) where East Asian text meets Latin letters or digits. The gap is a quarter em of the East Asian side, measured off Word's own PDF exports: an East Asian character standing before a Latin one advances 1.250 em against 1.000 em before another East Asian character, and a document that switches the feature off measures 1.000 em on both. Nothing is inserted next to a space, the two settings gate their own boundary, and both are parsed, honoured and written back.

## 0.1.0

### Minor Changes

- fb916eb: Release the DOCX and XLSX Python bindings as 0.1.0. They wrap the 0.2.0 engines and now share a version line with the PPTX binding.

### Patch Changes

- 15f0353: Ship the current DOCX engine in the Python binding: byte-stable round-trip saves that keep charts, opaque drawings and foreign markup, plus the pagination, table, list and font fidelity work since 0.0.2.
