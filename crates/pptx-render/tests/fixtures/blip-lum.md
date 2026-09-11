# Bitmap brightness and contrast

`blip-lum.pptx` is a synthetic, publishable reproduction of the missing `a:lum`.
Its only bitmap is a strip of six swatches — black, `404040`, `808080`,
`C0C0C0`, white and `03A7DF` — drawn five times at 288×48 from x=16, one row
every 60 px. The rows are the control, `bright="70000" contrast="-70000"`,
`bright="50000"`, `contrast="-50000"` and `bright="3000" contrast="77000"`. The
first and last pairs are the two that occur in the corpus.

Neither ECMA-376 nor MS-OI29500 gives a curve for `a:lum`, so this one is fitted
to LibreOffice 26.8, sampled over 121 `bright` × `contrast` pairs on a 16-step
grey ramp. With `b` and `c` as fractions:

```
slope  = c >= 0 ? 128 / (128 - 127c) : (128 + 127c) / 128
offset = 128 - 128 * slope + 255b * (1 + slope) / 2
out    = clamp(round(slope * in + offset), 0, 255)
```

applied per channel, alpha untouched. The reference agrees to within one level
on every pair sampled except `bright="70000" contrast="-70000"` exactly, which
LibreOffice diverts to its own watermark colour mode — its neighbours
(69 %/−70 %, 71 %/−70 %, 70 %/−69 %, 70 %/−71 %) all follow the formula above.
That shim is LibreOffice's, so it is not reproduced here; it leaves the washout
row 10–11 levels darker than the reference and nothing else moves.

| Sample at scale 1 | Source | Main | Fixed | LibreOffice |
| --- | --- | --- | --- | --- |
| control (40,40) | `000000` | `000000` | `000000` | `000000` |
| control (280,40) | `03A7DF` | `03A7DF` | `03A7DF` | `03A7DF` |
| washout (40,100) | `000000` | `000000` | `CDCDCD` | `D8D8D8` |
| washout (88,100) | `404040` | `404040` | `E1E1E1` | `EBEBEB` |
| washout (280,100) | `03A7DF` | `03A7DF` | `CEFFFF` | `D9FFFF` |
| bright 50 % (40,160) | `000000` | `000000` | `808080` | `7F7F7F` |
| contrast −50 % (40,220) | `000000` | `000000` | `404040` | `3F3F3F` |
| contrast −50 % (232,220) | `FFFFFF` | `FFFFFF` | `C0C0C0` | `BFBFBF` |
| contrast −50 % (280,220) | `03A7DF` | `03A7DF` | `4194B0` | `4193AF` |
| bright 3 %, contrast 77 % (136,280) | `808080` | `808080` | `949494` | `949494` |

`blip-lum.before.png` was rendered on main `1f30ea03`; the after image uses this
implementation.

Each swatch is eight source pixels wide so that a cell centre is exact under any
image filter. Sampling a one-pixel-per-swatch strip is not: both backends
interpolate the whole strip and no sample is then a pure swatch.
