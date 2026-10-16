`theme-fonts.pptx` is a public, synthetic two-slide repro for PR #291.

Slide 1 uses Liberation Serif as its major Latin font and Liberation Sans as its minor Latin font. Objects 2 and 3 request `+mj-lt`; object 4 requests `+mn-lt`; object 5 explicitly requests Arial. Slide 2 contains literal Light, Semibold, and Display family names plus an explicitly registered family as controls.

The widths below assume these registered families:

| Registered family | Asset prefix under `packages/fonts/assets/` |
| --- | --- |
| Arial | LiberationSans |
| Liberation Sans | LiberationSans |
| Liberation Serif | LiberationSerif |
| Courier New | LiberationMono |
| Calibri | Carlito |
| Times New Roman | LiberationSerif |

The two major-font text boxes use the major face. The heading is 437.0625 px wide and `#17365D`. The 400 px body box occupies three lines, with these breaks:

| Line |
| --- |
| `Theme major font chooses the ` |
| `heading face and its own ` |
| `wrapping metrics.` |

The minor text is `#008080`; the explicit Arial text is `#7F3F00`. Slide 2's first three line widths are 690.15625, 754.171875, and 725.6875 px.

In `shape-style.pptx`, slide 5, object 2 inherits a title style that requests `+mj-lt`. The theme's major font is Calibri Light, which is absent from the registry above, so the text uses the configured Arial/Liberation Sans Bold fallback (281 px), not Calibri/Carlito Bold. The text is `layout title 0070C0`, colored `#0070C0`. Registering the major face, as the public fixture does, selects that face directly.

Regression tests: `cargo test -p betteroffice-drawingml theme::tests` and `cargo test -p betteroffice-pptx-render --test theme_fonts`.
