`list-style-bullets.pptx` is a synthetic, one-slide repro for PR #294. It contains no private corpus material.

- The title placeholder inherits 66 pt, bold, centered, `#24265D` from its layout's `lstStyle`. The master fallback is 39.25 pt, regular, `#505050`.
- The body inherits `•` and `–` at levels 1 and 2. Level 3 is undefined in the list styles and retains the master's 14 pt gray text without a marker.
- The first bullet explicitly uses Courier New, `#D02020`, and 50% of its 18 pt text size. Its gutter starts at x=64 px; the next level's gutter starts at x=100 px.
- The last body paragraph resets bullet font, color, and size to follow its bold, green text.
- The second column exercises `defPPr`, a level-specific override, and a direct paragraph `buNone`/text-format override.

The display-list tests compare story ranges, line widths, caret stops, and hit tests against the same package with markers suppressed. Both point and percentage marker sizes are checked.

Bullet sizes and fonts follow the [DrawingML bullet size](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bulletsizepercentage) and [bullet font](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bulletfont) definitions.
