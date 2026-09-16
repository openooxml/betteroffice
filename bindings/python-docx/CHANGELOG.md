# @betteroffice/python-docx

## 0.1.0

### Minor Changes

- fb916eb: Release the DOCX and XLSX Python bindings as 0.1.0. They wrap the 0.2.0 engines and now share a version line with the PPTX binding.

### Patch Changes

- 15f0353: Ship the current DOCX engine in the Python binding: byte-stable round-trip saves that keep charts, opaque drawings and foreign markup, plus the pagination, table, list and font fidelity work since 0.0.2.
