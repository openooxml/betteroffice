# Comparison output

For DOCX and XLSX, `--png` writes the Vello render here alongside the same page
rendered by the existing raster backend, then prints per-channel mean absolute
difference and the share of pixels differing by more than 8 on any channel.
PPTX writes only the Vello image because it has no raster backend; its JSON
primitive summary and positioned-glyph audit are the verification artifact.

The images are generated, not committed. From the repository root:

```sh
cargo run --release --manifest-path apps/native-viewer/Cargo.toml -- \
  --png apps/native-viewer/artifacts/demo-page-1-vello.png

cargo run --release --manifest-path apps/native-viewer/Cargo.toml -- \
  --document apps/demo/public/showcase.xlsx --sheet 1 \
  --png apps/native-viewer/artifacts/showcase-sheet-1-vello.png

cargo run --release --manifest-path apps/native-viewer/Cargo.toml -- \
  --document apps/demo/public/betteroffice-demo.pptx --slide 2 \
  --png apps/native-viewer/artifacts/demo-slide-2-vello.png
```
