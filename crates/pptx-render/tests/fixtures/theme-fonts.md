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

The public fixture changes only the two major-font text boxes. The heading's width changes from 482.8828125 px to 437.0625 px. Its color remains `#17365D`. The 400 px body box still occupies three lines, with these breaks:

| Main | Fixed |
| --- | --- |
| `Theme major font chooses ` | `Theme major font chooses the ` |
| `the heading face and its ` | `heading face and its own ` |
| `own wrapping metrics.` | `wrapping metrics.` |

The minor text remains `#008080`; the explicit Arial text remains `#7F3F00`. Slide 2's first three line widths remain 690.15625, 754.171875, and 725.6875 px.

The only other changed slide is `shape-style.pptx`, slide 5, object 2. Its inherited title style requests `+mj-lt`. The theme's major font is Calibri Light, which is absent from the registry above. Main incorrectly uses Calibri/Carlito Bold (249.65625 px); the fixed build uses the configured Arial/Liberation Sans Bold fallback (281 px). The text remains `layout title 0070C0`, colored `#0070C0`. Registering the major face, as the public fixture does, selects that face directly.

Regression tests: `cargo test -p betteroffice-drawingml theme::tests` and `cargo test -p betteroffice-pptx-render --test theme_fonts`.
