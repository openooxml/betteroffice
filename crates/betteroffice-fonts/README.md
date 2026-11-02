# betteroffice-fonts

Decodes one font pack face using `zstd` or `woff2t+zstd`, with an optional
raw-content dictionary. Every BetterOffice engine links this decoder.
`decode_font_face` passes non-zstd input through and returns owned font bytes
plus the transformed payload for decoded woff2t faces.

The default output limit is 64 MiB. Explicit limits range from 0 to
2,147,483,647 bytes; woff2t decoding always caps the frame, dictionary, payload
and reconstructed font at 64 MiB. Frames require a content size and checksum.

Native targets decode zstd with the vendored zstd C decoder; wasm32 targets use
`ruzstd`, so building for wasm needs no C compiler with a wasm target. Both report
the same errors.

The `build` feature enables the `fontpack` encoder binary and its tooling
dependencies. The `wasm` feature builds the standalone decoder ABI.
