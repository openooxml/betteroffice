`shape-style.pptx` is a synthetic ten-slide colour-cascade fixture.

| Slide | Text colour source | Expected colour |
| --- | --- | --- |
| 1 | `fontRef` / theme 1 `lt1` | `#EEEEEE` |
| 2 | `fontRef` / `srgbClr` | `#FF0000` |
| 3 | Explicit run | `#00B050` |
| 4 | Paragraph `defRPr` | `#0070C0` |
| 5 | Layout title placeholder | `#0070C0` |
| 6 | Colourless `fontRef`, master `otherStyle` | `#595959` |
| 7 | Absent shape style, master `otherStyle` | `#595959` |
| 8 | Placeholder without inherited colour, `fontRef` | `#EEEEEE` |
| 9 | Slide → layout 2 → master 2 → theme 2 | `#123456` slide, `#234567` layout, `#345678` master |
| 10 | Master placeholder, paragraph default, explicit run | `#7030A0`, `#0070C0`, `#00B050` |

Slides 1, 2, 8, and 9 exercise the rendering change. All other slides retain
their colours from main. The fixture covers theme selection through relationships;
custom `clrMap` mappings and `themeOverride` parts are outside this change.
