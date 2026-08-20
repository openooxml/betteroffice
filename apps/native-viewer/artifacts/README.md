# Demo comparison

`demo-page-1-vello.png` is the GPU/Vello render. `demo-page-1-vello.raster.png` is the same display-list page rendered by `betteroffice-docx-raster`.

Generated from the repository root with:

```sh
cargo run --release --manifest-path apps/native-viewer/Cargo.toml -- --png apps/native-viewer/artifacts/demo-page-1-vello.png
```

The comparison uses threshold 8 on any RGBA channel. The exact metrics printed by the generating run are recorded in the implementation report.
