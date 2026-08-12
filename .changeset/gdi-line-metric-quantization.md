---
"@betteroffice/rust-crates": patch
---

Add GDI-compatible line-metric quantization to the text engine, and harden the single-spacing line box against malformed font tables.

Word's line boxes come from a text stack whose per-font vertical metrics are whole pixels at an integer ppem, while this engine scaled design metrics in pure floating point. The resulting sub-pixel error per line has the same sign for every line set in the same font and size, so it accumulates down a page until a line crosses a page boundary and pagination diverges. `single_line_box` can now reproduce the integer arithmetic: the em size rounds to a whole ppem, the metric family follows `OS/2` fsSelection bit 7 (USE_TYPO_METRICS, honored only from table version 4, since earlier versions reserve the bit), and ascent, descent and external leading each round to whole pixels before they are summed — rounding the parts, not the total.

This rides on a new `CompatFlags::gdi_line_metrics` and is **off by default**, so no document measures differently yet. Whether Word gates the behavior on `w:compatibilityMode` is not established, and the switch exists so a corpus harness can measure both paths before the default moves. `FontMetrics` gained `os2_fs_selection` and `os2_version` to support it.

The input guards around that function also tightened, on both the old and new paths. A font size of positive infinity previously passed the degenerate-input check and scaled into an infinite line box; it now yields the same all-zero box as NaN and non-positive sizes. Design metrics are clamped to a sane multiple of the em and the em size to Word's own 1638pt ceiling, so a font declaring a tiny `unitsPerEm` beside extreme `usWin*` values can no longer turn an ordinary font size into a line box the height of a document. Font bytes are attacker-controlled, so these are bounds, not style.
