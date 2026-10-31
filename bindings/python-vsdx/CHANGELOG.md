# @betteroffice/python-vsdx

## 0.0.2

### Patch Changes

- ad45f86: Adds opt-in `decode_font_face` calls for zstd and woff2t font pack faces. Rust facades expose the shared `betteroffice-fonts` decoder as `fonts`; Python returns font bytes and an optional transformed payload.
