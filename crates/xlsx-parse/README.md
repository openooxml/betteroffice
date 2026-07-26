# betteroffice-xlsx-parse

Streaming SpreadsheetML parser and serializer over
[betteroffice-xlsx-model](https://crates.io/crates/betteroffice-xlsx-model).

`parse_workbook` treats every byte as attacker-controlled: bounded nesting and
counts, and no allocation sized from a value in the file.

| Limit | Value |
| --- | --- |
| `MAX_DEPTH` | 64 |
| `MAX_CELLS` | 10,000,000 per worksheet |
| `MAX_SHARED_STRINGS` | 10,000,000 |
| `MAX_DEFINED_NAMES` | 65,536 |
| `MAX_HYPERLINKS` | 65,536 per worksheet |
| `MAX_STYLE_ENTRIES` | 65,536 |

`serialize_workbook` writes the model back out. It regenerates the parts the
model represents; package parts the model does not cover are not retained.

Used by [betteroffice-xlsx](https://crates.io/crates/betteroffice-xlsx).

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
