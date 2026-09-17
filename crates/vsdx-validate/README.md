# betteroffice-vsdx-validate

Read-only diagram rules over a parsed VSDX package: dangling connectors,
isolated shapes, overlapping shapes, crossing connectors and empty required
shape-data rows. `validate_package` walks every page, `validate_page` one, and
both return a deterministic, ordered list of issues. Nothing here mutates the
package, so the rules never touch the save projection.

Pages beyond `MAX_PAIRWISE_SHAPES` skip the two pairwise rules.

Part of [BetterOffice](https://github.com/openooxml/betteroffice).
