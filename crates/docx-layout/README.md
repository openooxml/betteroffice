# betteroffice-docx-layout

The DOCX pagination core. The contract is
`(MeasuredBlock[], LayoutOptions) -> Layout`: measurement happens in
[betteroffice-ooxml-text](https://crates.io/crates/betteroffice-ooxml-text),
this crate decides where everything lands on the page.

Covers paragraph flow with spacing collapse, page and column splits, explicit
breaks, multi-column options, inline and anchored images, text boxes, floating
objects, table cell layout, column balancing, keep-together and break policy,
header/footer bands, per-page footnote reservations, and hit testing. It emits
both the typed layout and a display list.

`layout_to_json` and `layout_to_canonical_json` are the pure entry points, unit
tested natively; `layout_document_json` is the `wasm_bindgen` wrapper. Layout
never touches the DOM — pages are replayed onto canvas by the host.

Used by [betteroffice-docx](https://crates.io/crates/betteroffice-docx).

Part of [BetterOffice](https://betteroffice.dev). Apache-2.0.
