# Comparison output

`--png` writes the Vello render here alongside the same page rendered by the
existing raster backend, then prints per-channel mean absolute difference and
the share of pixels differing by more than 8 on any channel.

The images are generated, not committed. From the repository root:

```sh
cargo run --release --manifest-path apps/native-viewer/Cargo.toml -- \
  --png apps/native-viewer/artifacts/demo-page-1-vello.png

cargo run --release --manifest-path apps/native-viewer/Cargo.toml -- \
  --document apps/demo/public/showcase.xlsx --sheet 1 \
  --png apps/native-viewer/artifacts/showcase-sheet-1-vello.png
```
