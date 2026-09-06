# Bitmap colour effects

`blip-effects.pptx` is a synthetic, publishable reproduction of #311. Its only
bitmap has six swatches: opaque `03A7DF`, black, white and red, followed by white
at alpha 128 and 0. All slide, layout and master pictures share that bitmap.

The six rows at x=16 are: control, biLevel at 50%, duotone, colour change followed
by duotone, colour change with `useA="0"`, and grayscale. The pictures at x=336
exercise the layout and master paths. The background is `D0D0D0`.

Duotone uses `bg2` (white `lt2`) shaded to `737373`, then white. Colour change
maps white to transparent red. `useA="0"` instead changes RGB and preserves the
original alpha. The default matches both RGB and alpha, keeping translucent
white edges intact. See [MS-OI29500, §2.1.1287](https://officeprotocoldoc.z19.web.core.windows.net/files/MS-OI29500/%5BMS-OI29500%5D-170919.pdf), page 506.

| Sample at scale 1 | Main | Fixed |
| --- | --- | --- |
| Control (40,40) | `03A7DF` | `03A7DF` |
| biLevel (40,100) | `03A7DF` | `000000` |
| duotone (40,160) | `03A7DF` | `B7B7B7` |
| knockout (136,220) | `FFFFFF` | `D0D0D0` background |
| useA=0 (136,280) | `FFFFFF` | `FF0000` |
| useA=0, half alpha (232,280) | `E8E8E8` | `E86868` |
| grayscale (40,340) | `03A7DF` | `7C7C7C` |

`blip-effects.before.png` was rendered on main `2c90c17f`; the after image uses
this implementation. The display-list comparison has eight images, seven with
effects added and one unchanged control. No other primitive fields change.

`crates/pptx-edit/tests/fixtures/blip-effects-main-v9.update.bin` was seeded from
this file by main `2c90c17f` with client ID 312. It covers schema v17 migration,
source attachment, save fidelity, and preservation of edits.

`blip-effects-main-v10.update.bin` was seeded by main `069e4d66`, and
`blip-effects-main-v16.update.bin` by main `49ecafba` with PR #295
(`dsaad68/pr/pptx-text-overflow` at `394052a4`, schema 16). The blip migration
is schema 17 because main took v10 through v16 for run baselines, numbering,
line spacing, picture fills, gradient outlines, chart properties and text
overflow. All three updates are tested with source attachment and an existing
edit.
