---
"@betteroffice/rust-crates": patch
---

Redaction now removes custom XML data and DOCX bindings by default while keeping redacted cached content. Custom DOCX and XLSX style names are anonymized, undefined DOCX style references are neutralized, and `RedactionReport::custom_xml_parts` counts removed parts.
