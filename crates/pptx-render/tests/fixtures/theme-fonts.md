`theme-fonts.pptx` is a public, synthetic two-slide repro for PR #291. The contributor's original deck is private; this fixture contains no private content.

Slide 1 uses Liberation Serif as its major Latin font and Liberation Sans as its minor Latin font. Objects 2 and 3 request `+mj-lt`; object 4 requests `+mn-lt`; object 5 explicitly requests Arial. Slide 2 contains literal Light, Semibold, and Display family names plus an explicitly registered family as controls.

Comparison baseline: `origin/main` at `2c90c17f5f5d871632cbafe680fb4a85a2a5d4d6`. Both builds register the same checked-in regular, bold, italic, and bold-italic font bytes, in that order, for these families:

| Registered family | Asset prefix under `packages/fonts/assets/` |
| --- | --- |
| Arial | LiberationSans |
| Liberation Sans | LiberationSans |
| Liberation Serif | LiberationSerif |
| Courier New | LiberationMono |
| Calibri | Carlito |
| Times New Roman | LiberationSerif |

The public fixture changes only the two major-font text boxes. The heading's width changes from 482.8828125 px to 437.0625 px. Its color remains `#17365D`. The 400 px body box still occupies three lines, with these breaks:

| Main | Fixed |
| --- | --- |
| `Theme major font chooses ` | `Theme major font chooses the ` |
| `the heading face and its ` | `heading face and its own ` |
| `own wrapping metrics.` | `wrapping metrics.` |

The minor text remains `#008080`; the explicit Arial text remains `#7F3F00`. Slide 2 is byte-identical. Its first three line widths remain 690.15625, 754.171875, and 725.6875 px. The contributor's suffix substitution instead produced 647.75, 741.625, and 677.96875 px; that unrelated substitution has been removed.

The only other changed slide is `shape-style.pptx`, slide 5, object 2. Its inherited title style requests `+mj-lt`. The theme's major font is Calibri Light, which is absent from the registry above. Main incorrectly uses Calibri/Carlito Bold (249.65625 px); the fixed build uses the configured Arial/Liberation Sans Bold fallback (281 px). The text remains `layout title 0070C0`, colored `#0070C0`. Registering the major face, as the public fixture does, selects that face directly.

Every tracked PPTX fixture was opened through `DeckSession::open`, and every slide's `SurfaceDisplayList` was serialized with `serde_json::to_vec`. Across the 20 existing decks, 56 of 57 slides are byte-identical; with the new fixture, 57 of 59 slides are byte-identical. On the two changed slides, only the affected paragraphs' font families and positioned text lines differ. All text content, colors, geometry, backgrounds, images, shape bounds, ordering, and other primitives match.

| Fixture | Slides | Changed slides (1-based) |
| --- | --- | --- |
| `apps/demo/public/betteroffice-demo.pptx` | 3 | None |
| `crates/ooxml-drawingml/tests/fixtures/preset-adjustments.pptx` | 3 | None |
| `crates/ooxml-opc/tests/fixtures/betteroffice-demo.pptx` | 3 | None |
| `crates/pptx-edit/tests/fixtures/deck-schema-v2-connectors.pptx` | 1 | None |
| `crates/pptx-edit/tests/fixtures/deck-schema-v2-nested-connectors.pptx` | 1 | None |
| `crates/pptx-edit/tests/fixtures/hidden-shapes.pptx` | 3 | None |
| `crates/pptx-edit/tests/fixtures/modern-comments.pptx` | 3 | None |
| `crates/pptx-parse/tests/fixtures/chart-deck.pptx` | 2 | None |
| `crates/pptx-parse/tests/fixtures/custom-geometry.pptx` | 2 | None |
| `crates/pptx-parse/tests/fixtures/master-style-deck.pptx` | 1 | None |
| `crates/pptx-parse/tests/fixtures/run-gradfill.pptx` | 2 | None |
| `crates/pptx-parse/tests/fixtures/shape-style.pptx` | 10 | 5 |
| `crates/pptx-parse/tests/fixtures/slide-number-fields.pptx` | 3 | None |
| `crates/pptx-parse/tests/fixtures/style-matrix-deck.pptx` | 2 | None |
| `crates/pptx-render/tests/fixtures/data-labels.pptx` | 5 | None |
| `crates/pptx-render/tests/fixtures/gradient-stop-order.pptx` | 3 | None |
| `crates/pptx-render/tests/fixtures/line-ends.pptx` | 3 | None |
| `crates/pptx-render/tests/fixtures/list-style-bullets.pptx` | 1 | None |
| `crates/pptx-render/tests/fixtures/picture-crop-mask.pptx` | 3 | None |
| `crates/pptx-render/tests/fixtures/run-paint-scope.pptx` | 3 | None |
| `crates/pptx-render/tests/fixtures/theme-fonts.pptx` | 2 | 1 |

All 21 deck snapshots are byte-identical. Each build deserialized and reserialized the other build's snapshots, and opened the other build's collaborative updates with the original source attached. No schema or writer changes are needed. A separate Python `zipfile` process compared no-edit saves from both builds against their inputs: all 342 file parts are byte-identical in each build. ZIP directory entries are not OPC parts and were excluded.

Regression tests: `cargo test -p betteroffice-drawingml theme::tests` and `cargo test -p betteroffice-pptx-render --test theme_fonts`.

Mutation checks cover every added or retained PR test: invert the major/minor selector; route East Asian and complex-script references to Latin; remove the empty DrawingML slot fallback; extend that fallback to WordprocessingML; restore the contributor's literal-name substitution. Each applicable test fails with its mutation and passes after restoration. The existing contributor resolver tests now use distinct, populated script slots so a missing script dispatch cannot pass accidentally.
