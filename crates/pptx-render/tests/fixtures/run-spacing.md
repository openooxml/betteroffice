# Run character spacing

`run-spacing.pptx` reduces the XML from issue #324 to one slide, using the tracked demo deck as its package shell. It contains the `OUR GREEN INITIATIVES` title from the issue, mixed positive/zero/negative tracking, a tracked `spAutoFit` body, and a shape list-style default. The original `green-solutions` and `swot-analysis` decks are in [document 4.zip](https://github.com/Suparnapaul393/PowerPoint-Sample/blob/master/document%204.zip), linked from the contributor’s evidence README.

Compared with main `069e4d66` using the same Liberation Sans font files:

| Shape | Main | Fixed |
| --- | --- | --- |
| Green title, `spc="600"`, 32 pt | 511.27078 px wide | 671.27075 px wide; 8 px gaps |
| Mixed tracking | One untracked run, 385.625 px | Three runs at 0, 8, and −1.33333 px; 410.9583 px |
| Shape autofit body | Shrunk to 21.33333 px | Keeps 42.66667 px and 4 px tracking |
| Inherited tracking | 461.47916 px | 529.4792 px with 4 px tracking |

The title stays centered at x=640 and retains `#008044`. The inherited text retains `#802020`, the mixed runs retain `#112233`, and the body retains `#224466`. Shape autofit no longer shrinks tracked text; growing its stored box remains a separate autofit feature.

`run-spacing.before.png` and `run-spacing.after.png` were rendered with the Rust raster backend. Tests cover parsing, tracking, wrapping, hard breaks, ligatures, mixed runs, autofit, editing, and source recovery from the committed main v10 update.

The original green-solutions title (object 29) also grows from 511.27078 to 671.27075 px while retaining its x=640 center and white `#FFFFFF` paint. Only that text primitive changes. Its screenshots are `run-spacing-green.before.png` and `run-spacing-green.after.png`.

On swot-analysis, only objects 42–49 change. STRENGTH grows from 132 to 160 px with 4 px gaps and retains `#0094C8`; the four body blocks keep their 16 px font and `#595959` colour while wrapping from five to seven lines. All other primitives are identical. These measurements verify tracking, not the remaining reference-renderer differences in shape autosizing or font metrics.
