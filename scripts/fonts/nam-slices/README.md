# Google Fonts CJK slicing strategies

`japanese_default.txt`, `korean_default.txt`, `simplified-chinese_default.txt`
and `traditional-chinese_default.txt` are copied unmodified from
[googlefonts/nam-files](https://github.com/googlefonts/nam-files) `slices/` at
commit `2a68014b16056e965c18fd47de1f3fde9a5b0095`, under the Apache License 2.0
in `LICENSE`. Each lists Unicode subsets from lowest to highest priority, ranked
by character frequency. `scripts/fonts/cjk-slices.py` reads them to order the
outline chunks of the CJK faces; they are build inputs and are not shipped.
