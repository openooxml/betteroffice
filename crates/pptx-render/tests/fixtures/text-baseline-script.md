`text-baseline-script.pptx` changes the demo deck's first-slide subtitle to
`E = mc² and H₂O are inline scripts.¹`, using baseline attributes rather than
Unicode script digits. The paragraph has an explicit zero default; its raised
runs use `30000` and its lowered run uses `-25000`, all at 17 pt.

Measurements at 96 DPI, using the tracked Liberation Sans regular, bold,
italic, and bold italic fonts registered as Arial:

| Run | Main `2c90c17f` size / baseline y | Fixed size / baseline y |
| --- | --- | --- |
| Plain text | 22.6667 px / 374.1195 px | 22.6667 px / 374.1195 px |
| Raised `2` and `1` | 22.6667 px / 374.1195 px | 13.1467 px / 367.3195 px |
| Lowered `2` | 22.6667 px / 374.1195 px | 13.1467 px / 379.7862 px |

Every run remains `#475467`. The text box is object 5,
`slide:0:256:shape:3`, at `(80, 344, 600, 112)` px. The before and after PNGs
use the same fonts and renderer settings. Of the 60 slides in all 21 tracked
PPTX fixtures, 59 display lists are byte-identical; only this subtitle changes.

`cargo test -p betteroffice-pptx-render baseline` checks the fixture, explicit
zero overrides, inherited shifts, adjacent opposite shifts, and line metrics.
`cargo test -p betteroffice-pptx-edit --test run_baseline` covers schema v10,
source recovery from actual main v9 updates, edits, and default JSON omission.
