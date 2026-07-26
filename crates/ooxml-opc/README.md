# betteroffice-opc

The OPC (OOXML zip) trust boundary. Every DOCX, XLSX, and PPTX package in
BetterOffice is unzipped and rezipped through this crate before a parser sees
a byte of it.

The limits are enforced by construction, so a hostile package cannot exhaust
memory or carry out-of-tree part names into a parser:

- 512 MiB total inflated bytes, 5,000 entries
- absolute paths, drive letters, and `..` traversal rejected on both separators
- slash, case, and dot aliases collapse to one key, so a security-sensitive
  part cannot be smuggled in twice under different spellings

`unzip_parts` and `rezip_parts` hold the logic and are unit-tested natively;
the `wasm_bindgen` wrappers only marshal to and from JS `{ path: Uint8Array }`
objects.

`sanitize_package` re-emits a package with disallowed parts and XML constructs
removed. `sanitize_package_for_format` additionally detects the real format and
rejects a package whose claimed format does not match what it actually is.

Used by [betteroffice-docx](https://crates.io/crates/betteroffice-docx),
[betteroffice-xlsx](https://crates.io/crates/betteroffice-xlsx), and
[betteroffice-pptx](https://crates.io/crates/betteroffice-pptx).

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
